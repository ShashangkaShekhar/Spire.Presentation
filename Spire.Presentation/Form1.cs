using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Spire.Presentation
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //create PPT document
            Presentation presentation = new Presentation();
            //add new shape to PPT document           
            IAutoShape shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.Rectangle,
                                new RectangleF(0, 50, 200, 50));
            shape.ShapeStyle.LineColor.Color = Color.White;
            shape.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.None;
            //add text to shape
            shape.AppendTextFrame("Hello World!");
            //set the Font fill style of text  
            TextRange textRange = shape.TextFrame.TextRange;
            textRange.Fill.FillType = Drawing.FillFormatType.Solid;
            textRange.Fill.SolidColor.Color = Color.Black;
            textRange.LatinFont = new TextFont("Arial Black");
            //save the document
            presentation.SaveToFile("hello.pptx", FileFormat.Pptx2010);
            System.Diagnostics.Process.Start("hello.pptx");
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {

            //add a smartart
            Presentation pres = new Presentation();
            Spire.Presentation.Diagrams.ISmartArt sa = pres.Slides[0].Shapes.AppendSmartArt(20, 40, 300, 300, Spire.Presentation.Diagrams.SmartArtLayoutType.Gear);

            //set type and color of smartart
            sa.Style = Spire.Presentation.Diagrams.SmartArtStyleType.SubtleEffect;
            sa.ColorStyle = Spire.Presentation.Diagrams.SmartArtColorType.GradientLoopAccent3;

            //remove all shapes
            foreach (object a in sa.Nodes)
                sa.Nodes.RemoveNode(0);

            //add two custom shapes with text
            Spire.Presentation.Diagrams.ISmartArtNode node = sa.Nodes.AddNode();
            sa.Nodes[0].TextFrame.Text = "aa";
            node = sa.Nodes.AddNode();
            node.TextFrame.Text = "bb";
            node.TextFrame.TextRange.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
            node.TextFrame.TextRange.Fill.SolidColor.KnownColor = KnownColors.Black;

            //save and launch the file
            pres.SaveToFile("SmartArtTest1.pptx", FileFormat.Pptx2007);
            System.Diagnostics.Process.Start("SmartArtTest1.pptx");

        }

        private void button3_Click(object sender, EventArgs e)
        {
            //create PPT document
            Presentation presentation = new Presentation();
            //add new table to PPT

            Double[] widths = new double[] { 100, 100, 150, 100, 100 };
            Double[] heights = new double[] { 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15 };
            ITable table = presentation.Slides[0].Shapes.AppendTable(presentation.SlideSize.Size.Width / 2 - 275, 80, widths, heights);
            //set the style of table
            table.StylePreset = TableStylePreset.LightStyle1Accent2;

            String[,] dataStr = new String[,]{
                                {"Name",    "Capital",  "Continent",    "Area", "Population"},
                                {"Venezuela",   "Caracas",  "South America",    "912047",   "19700000"},
                                {"Bolivia", "La Paz",   "South America",    "1098575",  "7300000"},
                                {"Brazil",  "Brasilia", "South America",    "8511196",  "150400000"},
                                {"Canada",  "Ottawa",   "North America",    "9976147",  "26500000"},
                                {"Chile",   "Santiago", "South America",    "756943",   "13200000"},
                                {"Colombia",    "Bagota",   "South America",    "1138907",  "33000000"},
                                {"Cuba",    "Havana",   "North America",    "114524",   "10600000"},
                                {"Ecuador", "Quito",    "South America",    "455502",   "10600000"},
                                {"Paraguay",    "Asuncion","South America", "406576",   "4660000"},
                                {"Peru",    "Lima", "South America",    "1285215",  "21600000"},
                                {"Jamaica", "Kingston", "North America",    "11424",    "2500000"},
                                {"Mexico",  "Mexico City",  "North America",    "1967180",  "88600000"}
                                };
            for (int i = 0; i < 13; i++)
                for (int j = 0; j < 5; j++)
                {
                    //fill the table with data
                    table[j, i].TextFrame.Text = dataStr[i, j];

                    //set the Font
                    table[j, i].TextFrame.Paragraphs[0].TextRanges[0].LatinFont = new TextFont("Arial Narrow");
                }

            //set the alignment of the first row to Center
            for (int i = 0; i < 5; i++)
            {
                table[i, 0].TextFrame.Paragraphs[0].Alignment = TextAlignmentType.Center;
            }
            //save the document
            presentation.SaveToFile("table.pptx", FileFormat.Pptx2010);
            System.Diagnostics.Process.Start("table.pptx");

        }
    }
}
