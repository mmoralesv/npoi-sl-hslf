
namespace NPOI.Util
{
    using System;
    using System.IO;
    using System.Text;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text.RegularExpressions;
    using NPOI.Common.UserModel;
    using NPOI.HSLF.Record;
    using static NPOI.Util.GenericRecordUtil;
    using System.Collections;
    public class GenericRecordJsonWriter : ICloseable, IDisposable
    {
        static char[] t = new char[255];
        private static string TABS = new String('\t', 255);
        private static string ZEROS = "0000000000000000";
        private static Regex ESC_CHARS = new Regex(@"[\\\p{C}\\\\]");
        private static string NL = System.Environment.NewLine;

        //Arrays.Fill(t, '\t')

        /**
         * Handler method
         *
         * @param record the parent record, applied via instance method reference
         * @param name the name of the property
         * @param object the value of the property
         * @return {@code true}, if the element was handled and output produced,
         *   The provided methods can be overridden and a implementation can return {@code false},
         *   if the element hasn't been written to the stream
         */
        protected delegate bool GenericRecordHandler(GenericRecordJsonWriter writer, string name, object o);

        private static Dictionary<Type, GenericRecordHandler> handler = new Dictionary<Type, GenericRecordHandler>
        {
            {typeof(String), (GenericRecordJsonWriter writer, string a, object o)=>writer.PrintObject(a,o)},
            {typeof(Number), (GenericRecordJsonWriter writer, string a, object o)=>writer.PrintNumber(a,o)},
            {typeof(Boolean), (GenericRecordJsonWriter writer, string a, object o)=>writer.PrintBoolean(a,o)},
            {typeof(IList), (GenericRecordJsonWriter writer, string a, object o)=>writer.PrintList(a,o)},
            {typeof(GenericRecord), (GenericRecordJsonWriter writer, string a, object o)=>writer.PrintGenericRecord(a,o)},
            {typeof(AnnotatedFlag), (GenericRecordJsonWriter writer, string a, object o)=>writer.PrintAnnotatedFlag(a,o)},
            {typeof(byte[]), (GenericRecordJsonWriter writer, string a, object o)=>writer.PrintBytes(a,o)},
            //{typeof(Point2D), (GenericRecordJsonWriter writer, string a, object o)=>writer.PrintPoint(a,o)},
            //{typeof(Dimension2D), (GenericRecordJsonWriter writer, string a, object o)=>writer.PrintDimension(a,o)},
            //{typeof(Rectangle2D), (GenericRecordJsonWriter writer, string a, object o)=>writer.PrintRectangle(a,o)},
            //{typeof(Path2D), (GenericRecordJsonWriter writer, string a, object o)=>writer.PrintPath(a,o)},
            //{typeof(AffineTransform), (GenericRecordJsonWriter writer, string a, object o)=>writer.PrintAffineTransform(a,o)},
            //{typeof(Color), (GenericRecordJsonWriter writer, string a, object o)=>writer.PrintColor(a,o)},
            //{typeof(BufferedImage), (GenericRecordJsonWriter writer, string a, object o)=>writer.PrintImage(a,o)},
            {typeof(Array), (GenericRecordJsonWriter writer, string a, object o)=>writer.PrintArray(a,o)},
            {typeof(Object), (GenericRecordJsonWriter writer, string a, object o)=>writer.PrintObject(a,o)},
        };
        private static void Handler<T>(T c, GenericRecordHandler printer)
        {
            handler.Add(c.GetType(), printer);
        }

        protected AppendableWriter aw;
        protected StreamWriter fw;
        protected int indent = 0;
        protected bool withComments = true;
        protected int childIndex = 0;

        public GenericRecordJsonWriter(FileInfo fileName)
        {
            OutputStream os;
            if ("null".Equals(fileName.Name))
            {
                os = null;
            }
            else
            {
                using (FileStream file = new FileStream(fileName.FullName, FileMode.Open))
                {
                    byte[] bytes = new byte[file.Length];
                    file.Read(bytes, 0, (int)file.Length);
                    os = (OutputStream)new MemoryStream();
                    os.Write(bytes, 0, (int)file.Length);
                }
            }
            //aw = new AppendableWriter(new StringBuilder();
            fw = new StreamWriter(os, Encoding.UTF8);
        }

        public GenericRecordJsonWriter(StringBuilder buffer)
        {
            aw = new AppendableWriter(buffer);
            var byteArray = Encoding.UTF8.GetBytes(aw.ToString());
            var os = (OutputStream)new MemoryStream(byteArray, true);
            fw = new StreamWriter(os, Encoding.UTF8);
        }

        public static String Marshal(GenericRecord record)
        {
            return Marshal(record, true);
        }

        public static String Marshal(GenericRecord record, bool withComments)
        {
            StringBuilder sb = new StringBuilder();
            try
            {

                using (GenericRecordJsonWriter w = new GenericRecordJsonWriter(sb))
                {
                    w.SetWithComments(withComments);
                    w.Write(record);
                    return sb.ToString();
                }
            }
            catch (IOException e)
            {
                return "{}";
            }
        }

        public void SetWithComments(bool withComments)
        {
            this.withComments = withComments;
        }

        public void Close()
        {
            fw.Close();
        }

        protected String Tabs()
        {
            return TABS.Substring(0, Math.Min(indent, TABS.Length));
        }

        public void Write(GenericRecord record)
        {
            string tabs = Tabs();
            RecordTypes type = record.GetGenericRecordType();
            String recordName;
            if (type != null)
            {
                recordName = Enum.GetName(typeof(RecordTypes), type);
            }
            else
            {
                recordName = typeof(RecordTypes).Name;
            }

            fw.Write(tabs);
            fw.Write("{");
            if (withComments)
            {
                fw.Write("   /* ");
                fw.Write(recordName);
                if (childIndex > 0)
                {
                    fw.Write(" - index: ");
                    fw.Write(childIndex);
                }
                fw.Write(" */");
            }
            fw.WriteLine();

            bool hasProperties = WriteProperties(record);
            fw.WriteLine();

            WriteChildren(record, hasProperties);

            fw.Write(tabs);
            fw.Write("}");
        }

        protected bool WriteProperties(GenericRecord record)
        {
            IDictionary<string, Func<object>> prop = record.GetGenericProperties<object>();
            if (prop == null || !prop.Any())
            {
                return false;
            }

            int oldChildIndex = childIndex;
            childIndex = 0;
            long cnt = prop.Count(e => WriteProp(e.Key, e.Value));
            childIndex = oldChildIndex;

            return cnt > 0;
        }


        protected bool WriteChildren(GenericRecord record, bool hasProperties)
        {
            List<GenericRecord> list = (List<GenericRecord>)record.GetGenericChildren();
            if (list == null || !list.Any())
            {
                return false;
            }

            indent++;
            aw.SetHoldBack(Tabs() + (hasProperties ? ", " : "") + "\"children\": [" + NL);
            int oldChildIndex = childIndex;
            childIndex = 0;
            long cnt = list.Where(l => WriteValue(null, l) && ++childIndex > 0).Count();
            childIndex = oldChildIndex;
            aw.SetHoldBack(null);

            if (cnt > 0)
            {
                fw.WriteLine();
                fw.WriteLine(Tabs() + "]");
            }
            indent--;

            return cnt > 0;
        }

        public void WriteError(String errorMsg)
        {
            fw.Write("{ error: ");
            PrintObject("error", errorMsg);
            fw.Write(" }");
        }

        protected bool WriteProp(String name, Func<object> value)
        {
            bool isNext = (childIndex > 0);
            aw.SetHoldBack(isNext ? NL + Tabs() + "\t, " : Tabs() + "\t  ");
            int oldChildIndex = childIndex;
            childIndex = 0;
            bool written = WriteValue(name, value?.Invoke());
            childIndex = oldChildIndex + (written ? 1 : 0);
            aw.SetHoldBack(null);
            return written;
        }

        protected bool WriteValue(String name, Object o)
        {
            if (childIndex > 0)
            {
                aw.SetHoldBack(",");
            }

            GenericRecordHandler grh = null;
            if (o == null)
            {
                PrintNull(name, o);
            }
            else
            {
                grh = handler.FirstOrDefault(h => matchInstanceOrArray(h.Key, o)).Value;
            }

            bool result = grh != null && grh.Invoke(this, name, o);
            aw.SetHoldBack(null);
            return result;
        }

        protected bool matchInstanceOrArray(Type key, Object instance)
        {
            return instance.GetType().IsInstanceOfType(key) || (key.IsArray && (instance.GetType().IsInstanceOfType(key.GetElementType())));
        }

        protected void PrintName(String name)
        {
            fw.Write(name != null ? "\"" + name + "\": " : "");
        }

        protected bool PrintNull(String name, Object o)
        {
            PrintName(name);
            fw.Write("null");
            return true;
        }

        //@SuppressWarnings("java:S3516")
        protected bool PrintNumber(String name, Object o)
        {
            Number n = (Number)o;
            PrintName(name);

            if (o.GetType().IsAssignableFrom(typeof(float)))
            {
                fw.Write(n.GetFloatValue());
                return true;
            }
            else if (o.GetType().IsAssignableFrom(typeof(double)))
            {
                fw.Write(n.GetDoubleValue());
                return true;
            }

            fw.Write(n.GetLongValue());

            int size;
            if (n.GetType().IsAssignableFrom(typeof(byte)))
            {
                size = 2;
            }
            else if (n.GetType().IsAssignableFrom(typeof(short)))
            {
                size = 4;
            }
            else if (n.GetType().IsAssignableFrom(typeof(int)))
            {
                size = 8;
            }
            else if (n.GetType().IsAssignableFrom(typeof(long)))
            {
                size = 16;
            }
            else
            {
                size = -1;
            }

            long l = n.GetLongValue();
            if (withComments && size > 0 && (l < 0 || l > 9))
            {
                fw.Write(" /* 0x");
                fw.Write(TrimHex(l, size));
                fw.Write(" */");
            }
            return true;
        }

        protected bool PrintBoolean(String name, Object o)
        {
            PrintName(name);
            fw.Write(((bool)o).ToString());
            return true;
        }

        protected bool PrintList(String name, Object o)
        {
            PrintName(name);
            fw.WriteLine("[");
            int oldChildIndex = childIndex;
            childIndex = 0;
            foreach (var e in (IList)o)
            {
                WriteValue(null, e);
                childIndex++;
            }

            childIndex = oldChildIndex;
            fw.Write(Tabs() + "\t]");
            return true;
        }

        protected bool PrintGenericRecord(String name, Object o)
        {
            PrintName(name);
            this.indent++;
            Write((GenericRecord)o);
            this.indent--;
            return true;
        }

        protected bool PrintAnnotatedFlag(String name, Object o)
        {

            PrintName(name);
            AnnotatedFlag af = (AnnotatedFlag)o;
            fw.Write((long)af.GetValue().Invoke());
            if (withComments)
            {
                fw.Write(" /* ");
                fw.Write(af.GetDescription());
                fw.Write(" */ ");
            }
            return true;
        }

        protected bool PrintBytes(String name, Object o)
        {
            PrintName(name);
            fw.Write('"');
            fw.Write(Convert.ToBase64String((byte[])o));
            fw.Write('"');
            return true;
        }

        // protected bool PrintPoint(String name, Object o)
        // {
        //     PrintName(name);
        //     Point2D p = (Point2D)o;
        //     fw.Write("{ \"x\": " + p.getX() + ", \"y\": " + p.getY() + " }");
        //     return true;
        // }

        // protected bool PrintDimension(String name, Object o)
        // {
        //     PrintName(name);
        //     Dimension2D p = (Dimension2D)o;
        //     fw.Write("{ \"width\": " + p.getWidth() + ", \"height\": " + p.GetHeight() + " }");
        //     return true;
        // }

        //protected bool PrintRectangle(String name, Object o)
        //{
        //    printName(name);
        //    Rectangle2D p = (Rectangle2D)o;
        //    fw.Write("{ \"x\": " + p.getX() + ", \"y\": " + p.getY() + ", \"width\": " + p.getWidth() + ", \"height\": " + p.getHeight() + " }");
        //    return true;
        //}

        // protected bool PrintPath(String name, Object o)
        // {
        //     PrintName(name);
        //     PathIterator iter = ((Path2D)o).getPathIterator(null);
        //     double[] pnts = new double[6];
        //     fw.Write("[");
        //
        //     indent += 2;
        //     String t = tabs();
        //     indent -= 2;
        //
        //     bool isNext = false;
        //     while (!iter.isDone())
        //     {
        //         fw.WriteLine(isNext ? ", " : "");
        //         fw.Write(t);
        //         isNext = true;
        //         int segType = iter.currentSegment(pnts);
        //         fw.Write("{ \"type\": ");
        //         switch (segType)
        //         {
        //             case PathIterator.SEG_MOVETO:
        //                 fw.Write("\"move\", \"x\": " + pnts[0] + ", \"y\": " + pnts[1]);
        //                 break;
        //             case PathIterator.SEG_LINETO:
        //                 fw.Write("\"lineto\", \"x\": " + pnts[0] + ", \"y\": " + pnts[1]);
        //                 break;
        //             case PathIterator.SEG_QUADTO:
        //                 fw.Write("\"quad\", \"x1\": " + pnts[0] + ", \"y1\": " + pnts[1] + ", \"x2\": " + pnts[2] + ", \"y2\": " + pnts[3]);
        //                 break;
        //             case PathIterator.SEG_CUBICTO:
        //                 fw.Write("\"cubic\", \"x1\": " + pnts[0] + ", \"y1\": " + pnts[1] + ", \"x2\": " + pnts[2] + ", \"y2\": " + pnts[3] + ", \"x3\": " + pnts[4] + ", \"y3\": " + pnts[5]);
        //                 break;
        //             case PathIterator.SEG_CLOSE:
        //                 fw.Write("\"close\"");
        //                 break;
        //         }
        //         fw.Write(" }");
        //         iter.next();
        //     }
        //
        //     fw.Write("]");
        //     return true;
        // }

        protected bool PrintObject(String name, Object o)
        {
            PrintName(name);
            fw.Write('"');

            String str = o.ToString();
            var matcher = ESC_CHARS.Matches(str);
            int pos = 0;

            foreach (System.Text.RegularExpressions.Match m in matcher)
            {
                fw.Write(str, pos, m.Index);
                string match = m.Groups[1].Value;
                switch (match)
                {
                    case "\n":
                        fw.Write("\\\\n");
                        break;
                    case "\r":
                        fw.Write("\\\\r");
                        break;
                    case "\t":
                        fw.Write("\\\\t");
                        break;
                    case "\b":
                        fw.Write("\\\\b");
                        break;
                    case "\f":
                        fw.Write("\\\\f");
                        break;
                    case "\\":
                        fw.Write("\\\\\\\\");
                        break;
                    case "\"":
                        fw.Write("\\\\\"");
                        break;
                    default:
                        fw.Write("\\\\u");
                        fw.Write(TrimHex(match[0], 4));
                        break;
                }
                pos = m.Index + m.Length;
            }

            fw.Write(str, pos, str.Length);
            fw.Write('"');
            return true;
        }

        //protected bool PrintAffineTransform(String name, Object o)
        //{
        //    printName(name);
        //    AffineTransform xForm = (AffineTransform)o;
        //    fw.write(
        //        "{ \"scaleX\": " + xForm.getScaleX() +
        //        ", \"shearX\": " + xForm.getShearX() +
        //        ", \"transX\": " + xForm.getTranslateX() +
        //        ", \"scaleY\": " + xForm.getScaleY() +
        //        ", \"shearY\": " + xForm.getShearY() +
        //        ", \"transY\": " + xForm.getTranslateY() + " }");
        //    return true;
        //}

        //protected bool PrintColor(String name, Object o)
        //{
        //    printName(name);

        //    int rgb = ((Color)o).getRGB();
        //    fw.print(rgb);

        //    if (withComments)
        //    {
        //        fw.write(" /* 0x");
        //        fw.write(trimHex(rgb, 8));
        //        fw.write(" */");
        //    }
        //    return true;
        //}

        protected bool PrintArray(String name, Object o)
        {
            PrintName(name);
            fw.Write("[");
            int length = (o as Array).Length;

            int oldChildIndex = childIndex;
            for (childIndex = 0; childIndex < length; childIndex++)
            {
                WriteValue(null, (o as Array).GetValue(childIndex));
            }
            childIndex = oldChildIndex;
            fw.Write(Tabs() + "\t]");
            return true;
        }

        //protected bool PrintImage(String name, Object o)
        //{
        //    BufferedImage img = (Bitmap)o;
        //    BufferedImage img = (BufferedImage)o;


        //    String[] COLOR_SPACES = {
        //    "XYZ","Lab","Luv","YCbCr","Yxy","RGB","GRAY","HSV","HLS","CMYK","Unknown","CMY","Unknown"

        //};

        //    String[] IMAGE_TYPES = {
        //    "CUSTOM","INT_RGB","INT_ARGB","INT_ARGB_PRE","INT_BGR","3BYTE_BGR","4BYTE_ABGR","4BYTE_ABGR_PRE",
        //        "USHORT_565_RGB","USHORT_555_RGB","BYTE_GRAY","USHORT_GRAY","BYTE_BINARY","BYTE_INDEXED"

        //};

        //    PrintName(name);
        //    ColorModel cm = img.getColorModel();
        //    //String colorType =
        //    //    (cm instanceof IndexColorModel) ? "indexed" :
        //    //(cm instanceof ComponentColorModel) ? "component" :
        //    //(cm instanceof DirectColorModel) ? "direct" :
        //    //(cm instanceof PackedColorModel) ? "packed" : "unknown";

        //    if(cm.GetType() == typeof(IndexColorModel))
        //    {

        //    }

        //    fw.write(
        //        "{ \"width\": " + img.getWidth() +
        //        ", \"height\": " + img.getHeight() +
        //        ", \"type\": \"" + IMAGE_TYPES[img.getType()] + "\"" +
        //        ", \"colormodel\": \"" + colorType + "\"" +
        //        ", \"pixelBits\": " + cm.getPixelSize() +
        //        ", \"numComponents\": " + cm.getNumComponents() +
        //        ", \"colorSpace\": \"" + COLOR_SPACES[Math.Min(cm.getColorSpace().getType(), 12)] + "\"" +
        //        ", \"transparency\": " + cm.getTransparency() +
        //        ", \"alpha\": " + cm.hasAlpha() +
        //        "}"
        //    );
        //    return true;
        //}

        public static String TrimHex(long l, int size)
        {
            String b = Convert.ToString(l, 16);
            int len = b.Length;
            return ZEROS.Substring(0, Math.Max(0, size - len)) + b.Substring(Math.Max(0, len - size), len);
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (disposing)
            {
                aw = null;
                fw = null;
            }
        }
    }

    public class AppendableWriter : TextWriter
    {

        private StringBuilder appender;
        private TextWriter writer;
        private string holdBack;

        public override Encoding Encoding => throw new NotImplementedException();

        public AppendableWriter(StringBuilder buffer)
        {
            //super(buffer);
            this.appender = buffer;
            this.writer = null;
        }

        AppendableWriter(TextWriter writer)
        {
            //super(writer);
            this.appender = null;
            this.writer = writer;
        }

        protected internal void SetHoldBack(String holdBack)
        {
            this.holdBack = holdBack;
        }

        //@Override
        public override void Write(char[] cbuf, int off, int len)
        {
            if (holdBack != null)
            {
                if (appender != null)
                {
                    appender.Append(holdBack);
                }
                else if (writer != null)
                {
                    writer.Write(holdBack);
                }
                holdBack = null;
            }

            if (appender != null)
            {
                appender.Append(string.Join("", cbuf), off, len);
            }
            else if (writer != null)
            {
                writer.Write(cbuf, off, len);
            }
        }

        //@Override
        public override void Flush()
        {
            object o;
            if (appender != null)
            {
                o = appender;
            }
            else
            {
                o = writer;
            }
            if (o.GetType() == typeof(Stream))
            {
                ((Stream)o).Flush();
            }
        }

        //@Override
        public override void Close()
        {
            Flush();
            Object o;
            if (appender != null)
            {
                o = appender;
            }
            else
            {
                o = writer;
            };
            if (o is ICloseable)
            {
                ((ICloseable)o).Close();
            }
        }
    }
}

