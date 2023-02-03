using System.Linq;
using NPOI.DDF;
using NPOI.Util;

namespace NPOI.HSLF.UserModel
{
    using System.Collections.Generic;
    using NPOI.SL.UserModel;
    using SixLabors.ImageSharp;
    
    public class HSLFSimpleGradientPaint : IGradientPaint
    {
        private List<float> fractions;
        private List<Color> colors;
        private HSLFShape shape;
        
        public PaintModifier PaintModifier { get; }
        public FlipMode FlipMode { get; }
        public TextureAlignment TextureAlignment { get; }
        public double GradientAngle { get; }
        public ColorStyle[] GradientColors { get; }
        public float[] GradientFractions { get; }
        public GradientType GradientType1 { get; }
        public Insets2D FillToInsets { get; }
        
        public HSLFSimpleGradientPaint(List<float> fractions, List<Color> colors, HSLFShape shape)
        {
            this.fractions = fractions;
            this.colors = colors;
            this.shape = shape;
        }
        
        public bool IsRotatedWithShape()
        {
            throw new System.NotImplementedException();
        }
        
        public double GetGradientAngle()
        {
            // A value of type FixedPoint, as specified in [MS-OSHARED] section 2.2.1.6,
            // that specifies the angle of the gradient fill. Zero degrees represents a vertical vector from
            // bottom to top. The default value for this property is 0x00000000.
            int rot = shape.GetEscherProperty(EscherProperties.FILL__ANGLE);
            return 90-Units.FixedPointToDecimal(rot);
        }

        public ColorStyle[] GetGradientColors()
        {
            //return colors.map(this.WrapColor).toArray(ColorStyle[]::new);
            return colors.Select(color => WrapColor(color)).ToArray();
        }

        private ColorStyle WrapColor(Color col)
        {
                    
            //return (col == null) ? null : DrawPaint.createSolidPaint(col).getSolidColor();
            return new ColorStyle();
        }

        public float[] GetGradientFractions()
        {
            float[] frc = new float[fractions.Capacity];
            for (int i = 0; i < fractions.Capacity; i++)
            {
                frc[i] = fractions[i];
            }

            return frc;
        }

        // public bool IsRotatedWithShape()
        // {
        //     return HSLFFill.IsRotatedWithShape();
        // }

        public GradientType GetGradientType()
        {
            return GradientType1;
        }
    }
}