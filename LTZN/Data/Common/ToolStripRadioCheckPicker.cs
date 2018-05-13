using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using System.Drawing;

namespace LTZN.Common
{
    class ToolStripRadioButtonPicker: ToolStripControlHost
    {
        public ToolStripRadioButtonPicker() : base(new RadioButton()) { }

        public RadioButton RadioButtonControl
        {
            get
            {
                return Control as RadioButton;
            }
        }

        public string RadioText
        {
            get { return RadioButtonControl.Text; }
            set { RadioButtonControl.Text = value; }
        }
        public TextImageRelation RadioTextImageRelation
        {
            get { return RadioButtonControl.TextImageRelation; }
            set { RadioButtonControl.TextImageRelation = value; }
        }
        public Font RadioFont
        {
            get { return RadioButtonControl.Font; }
            set { RadioButtonControl.Font = value; }
        }
        public int RadioWidth
        {
            get { return RadioButtonControl.Width; }
            set { RadioButtonControl.Width = value; }
        }
        public int RadioHeight
        {
            get { return RadioButtonControl.Height; }
            set { RadioButtonControl.Height = value; }
        }
        public bool RadioChecked
        {
            get { return RadioButtonControl.Checked; }
            set { RadioButtonControl.Checked = value; }
        }
        public ContentAlignment RadioCheckAlign
        {
            get { return RadioButtonControl.CheckAlign; }
            set { RadioButtonControl.CheckAlign = value; }
        }
        // Subscribe and unsubscribe the control events you wish to expose.
        protected override void OnSubscribeControlEvents(Control c)
        {
            // Call the base so the base events are connected.
            base.OnSubscribeControlEvents(c);

            // Cast the control to a RadioButton control.
            RadioButton RadioButtonControl = (RadioButton)c;

            // Add the event.
            RadioButtonControl.CheckedChanged += new EventHandler(RadioButtonCheckedChanged);

        }
        protected override void OnUnsubscribeControlEvents(Control c)
        {
            // Call the base method so the basic events are unsubscribed.
            base.OnUnsubscribeControlEvents(c);

            // Cast the control to a RadioButton control.
            RadioButton RadioButtonControl = (RadioButton)c;

            // Remove the event.
            RadioButtonControl.CheckedChanged -=
                new EventHandler(RadioButtonCheckedChanged);
        }

        // 为要公开的事件提供必要的包装。
        public event EventHandler  CheckedChanged;
        public void RadioButtonCheckedChanged(object sender, EventArgs e)
        {
            if (CheckedChanged != null)
            {
                CheckedChanged(this, e);
            }
        }
    }
}
