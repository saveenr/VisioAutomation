namespace VisioAutomation.UI
{
    partial class ThreePointFillControl
    {
        /// <summary> 
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.radioButtonUp = new System.Windows.Forms.RadioButton();
            this.radioButtonRight = new System.Windows.Forms.RadioButton();
            this.radioButtonDown = new System.Windows.Forms.RadioButton();
            this.radioButtonLeft = new System.Windows.Forms.RadioButton();
            this.groupBoxDirection = new System.Windows.Forms.GroupBox();
            this.labelColorEdge = new System.Windows.Forms.Label();
            this.labelColorRight = new System.Windows.Forms.Label();
            this.labelColorLeft = new System.Windows.Forms.Label();
            this.ColorPickerPrimaryEdge = new VisioAutomation.UI.CommonControls.ColorSelectorSmall();
            this.ColorPickerCorner2 = new VisioAutomation.UI.CommonControls.ColorSelectorSmall();
            this.ColorPickerCorner1 = new VisioAutomation.UI.CommonControls.ColorSelectorSmall();
            this.buttonSwapCorner = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.groupBoxDirection.SuspendLayout();
            this.SuspendLayout();
            // 
            // radioButtonUp
            // 
            this.radioButtonUp.AutoSize = true;
            this.radioButtonUp.Location = new System.Drawing.Point(38, 22);
            this.radioButtonUp.Name = "radioButtonUp";
            this.radioButtonUp.Size = new System.Drawing.Size(39, 17);
            this.radioButtonUp.TabIndex = 3;
            this.radioButtonUp.TabStop = true;
            this.radioButtonUp.Text = "Up";
            this.radioButtonUp.UseVisualStyleBackColor = true;
            // 
            // radioButtonRight
            // 
            this.radioButtonRight.AutoSize = true;
            this.radioButtonRight.Location = new System.Drawing.Point(63, 45);
            this.radioButtonRight.Name = "radioButtonRight";
            this.radioButtonRight.Size = new System.Drawing.Size(45, 17);
            this.radioButtonRight.TabIndex = 4;
            this.radioButtonRight.TabStop = true;
            this.radioButtonRight.Text = "right";
            this.radioButtonRight.UseVisualStyleBackColor = true;
            // 
            // radioButtonDown
            // 
            this.radioButtonDown.AutoSize = true;
            this.radioButtonDown.Location = new System.Drawing.Point(38, 68);
            this.radioButtonDown.Name = "radioButtonDown";
            this.radioButtonDown.Size = new System.Drawing.Size(53, 17);
            this.radioButtonDown.TabIndex = 5;
            this.radioButtonDown.TabStop = true;
            this.radioButtonDown.Text = "Down";
            this.radioButtonDown.UseVisualStyleBackColor = true;
            // 
            // radioButtonLeft
            // 
            this.radioButtonLeft.AutoSize = true;
            this.radioButtonLeft.Location = new System.Drawing.Point(14, 45);
            this.radioButtonLeft.Name = "radioButtonLeft";
            this.radioButtonLeft.Size = new System.Drawing.Size(43, 17);
            this.radioButtonLeft.TabIndex = 6;
            this.radioButtonLeft.TabStop = true;
            this.radioButtonLeft.Text = "Left";
            this.radioButtonLeft.UseVisualStyleBackColor = true;
            // 
            // groupBoxDirection
            // 
            this.groupBoxDirection.Controls.Add(this.radioButtonLeft);
            this.groupBoxDirection.Controls.Add(this.radioButtonRight);
            this.groupBoxDirection.Controls.Add(this.radioButtonDown);
            this.groupBoxDirection.Controls.Add(this.radioButtonUp);
            this.groupBoxDirection.Location = new System.Drawing.Point(13, 132);
            this.groupBoxDirection.Name = "groupBoxDirection";
            this.groupBoxDirection.Size = new System.Drawing.Size(128, 99);
            this.groupBoxDirection.TabIndex = 7;
            this.groupBoxDirection.TabStop = false;
            this.groupBoxDirection.Text = "Direction from Edge";
            this.groupBoxDirection.Enter += new System.EventHandler(this.groupBoxDirection_Enter);
            // 
            // labelColorEdge
            // 
            this.labelColorEdge.AutoSize = true;
            this.labelColorEdge.Location = new System.Drawing.Point(52, 67);
            this.labelColorEdge.Name = "labelColorEdge";
            this.labelColorEdge.Size = new System.Drawing.Size(32, 13);
            this.labelColorEdge.TabIndex = 8;
            this.labelColorEdge.Text = "Edge";
            // 
            // labelColorRight
            // 
            this.labelColorRight.AutoSize = true;
            this.labelColorRight.Location = new System.Drawing.Point(106, 9);
            this.labelColorRight.Name = "labelColorRight";
            this.labelColorRight.Size = new System.Drawing.Size(32, 13);
            this.labelColorRight.TabIndex = 9;
            this.labelColorRight.Text = "Right";
            // 
            // labelColorLeft
            // 
            this.labelColorLeft.AutoSize = true;
            this.labelColorLeft.Location = new System.Drawing.Point(7, 9);
            this.labelColorLeft.Name = "labelColorLeft";
            this.labelColorLeft.Size = new System.Drawing.Size(25, 13);
            this.labelColorLeft.TabIndex = 10;
            this.labelColorLeft.Text = "Left";
            // 
            // ColorPickerPrimaryEdge
            // 
            this.ColorPickerPrimaryEdge.Color = System.Drawing.SystemColors.Control;
            this.ColorPickerPrimaryEdge.Location = new System.Drawing.Point(55, 83);
            this.ColorPickerPrimaryEdge.Name = "ColorPickerPrimaryEdge";
            this.ColorPickerPrimaryEdge.Size = new System.Drawing.Size(64, 22);
            this.ColorPickerPrimaryEdge.TabIndex = 2;
            // 
            // ColorPickerCorner2
            // 
            this.ColorPickerCorner2.Color = System.Drawing.SystemColors.Control;
            this.ColorPickerCorner2.Location = new System.Drawing.Point(109, 28);
            this.ColorPickerCorner2.Name = "ColorPickerCorner2";
            this.ColorPickerCorner2.Size = new System.Drawing.Size(42, 22);
            this.ColorPickerCorner2.TabIndex = 1;
            // 
            // ColorPickerCorner1
            // 
            this.ColorPickerCorner1.Color = System.Drawing.SystemColors.Control;
            this.ColorPickerCorner1.Location = new System.Drawing.Point(10, 28);
            this.ColorPickerCorner1.Name = "ColorPickerCorner1";
            this.ColorPickerCorner1.Size = new System.Drawing.Size(42, 22);
            this.ColorPickerCorner1.TabIndex = 0;
            // 
            // buttonSwapCorner
            // 
            this.buttonSwapCorner.Location = new System.Drawing.Point(147, 200);
            this.buttonSwapCorner.Name = "buttonSwapCorner";
            this.buttonSwapCorner.Size = new System.Drawing.Size(115, 23);
            this.buttonSwapCorner.TabIndex = 11;
            this.buttonSwapCorner.Text = "Swap Corners";
            this.buttonSwapCorner.UseVisualStyleBackColor = true;
            this.buttonSwapCorner.Click += new System.EventHandler(this.buttonSwapCorner_Click);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(147, 171);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(115, 23);
            this.button1.TabIndex = 12;
            this.button1.Text = "Rotate Colors";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // ThreePointFillControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.button1);
            this.Controls.Add(this.buttonSwapCorner);
            this.Controls.Add(this.labelColorLeft);
            this.Controls.Add(this.labelColorRight);
            this.Controls.Add(this.labelColorEdge);
            this.Controls.Add(this.ColorPickerPrimaryEdge);
            this.Controls.Add(this.ColorPickerCorner2);
            this.Controls.Add(this.ColorPickerCorner1);
            this.Controls.Add(this.groupBoxDirection);
            this.Name = "ThreePointFillControl";
            this.Size = new System.Drawing.Size(279, 241);
            this.Load += new System.EventHandler(this.UC3PointFill_Load);
            this.groupBoxDirection.ResumeLayout(false);
            this.groupBoxDirection.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private VisioAutomation.UI.CommonControls.ColorSelectorSmall ColorPickerCorner1;
        private VisioAutomation.UI.CommonControls.ColorSelectorSmall ColorPickerCorner2;
        private VisioAutomation.UI.CommonControls.ColorSelectorSmall ColorPickerPrimaryEdge;
        private System.Windows.Forms.RadioButton radioButtonUp;
        private System.Windows.Forms.RadioButton radioButtonRight;
        private System.Windows.Forms.RadioButton radioButtonDown;
        private System.Windows.Forms.RadioButton radioButtonLeft;
        private System.Windows.Forms.GroupBox groupBoxDirection;
        private System.Windows.Forms.Label labelColorEdge;
        private System.Windows.Forms.Label labelColorRight;
        private System.Windows.Forms.Label labelColorLeft;
        private System.Windows.Forms.Button buttonSwapCorner;
        private System.Windows.Forms.Button button1;
    }
}
