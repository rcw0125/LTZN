using System;
using System.Collections.Generic;
using System.Text;
using System.Drawing;

namespace LTZN.高炉燃料比综合分析
{
    class colors
    {
        public List<colors> Colors = new List<colors>();

        public static  Color getColor(int index)
        {

            Random rand1 = new Random(index * 11);
            Random rand2 = new Random(index * 17);
            Random rand3 = new Random(index * 29);
            return Color.FromArgb(rand1.Next(250), rand1.Next(250), rand1.Next(250));
        }
    }
}
