using System;
using System.Collections.Generic;
using System.Text;

namespace LTZN.CalFramework
{
    public class FormaCal : CalTagDecorator
    {
        /// <remarks>公式</remarks>
        public string Forma
        {
            get;
            set;
        }

        public override void Cal()
        {
            if (!CalFinished)
            {
                if (!string.IsNullOrEmpty(Forma))
                {
                    this.TagValue=CalService.DoForma(Forma);
                    CalFinished = true;
                }
            }
        }

        public FormaCal(CalTag tag)
            : base(tag)
        {

        }
    }
}
