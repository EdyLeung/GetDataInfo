using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GetDataInfo
{
    public partial class FrmSetConfig : Form
    {
       private CommonRegex comReg = new CommonRegex();
       private INIFileOperate iniInfo = new INIFileOperate();
       private string configIniPath = AppDomain.CurrentDomain.SetupInformation.ApplicationBase + "\\Config.ini";	//INI文件的物理地址;

        public FrmSetConfig()
        {
            InitializeComponent();

            comReg.bindCbox(this.cboName);
            comReg.bindCbox(this.cboIdCard);
            comReg.bindCbox(this.cboArea);
            comReg.bindCbox(this.cboBirthday);
            comReg.bindCbox(this.cboSex);

            this.cboName.SelectedIndexChanged += cboboBoxs_SelectedIndexChanged;
            this.cboIdCard.SelectedIndexChanged += cboboBoxs_SelectedIndexChanged;
            this.cboArea.SelectedIndexChanged += cboboBoxs_SelectedIndexChanged;
            this.cboBirthday.SelectedIndexChanged += cboboBoxs_SelectedIndexChanged;
            this.cboSex.SelectedIndexChanged += cboboBoxs_SelectedIndexChanged;
            //regexIniPath = AppDomain.CurrentDomain.SetupInformation.ApplicationBase + "\\RegexString.ini";	//INI文件的物理地址;
            //picIniPath = AppDomain.CurrentDomain.SetupInformation.ApplicationBase + "\\PicSize.ini";	//INI文件的物理地址;
            LoadSeting();
        }

        void cboboBoxs_SelectedIndexChanged(object sender, EventArgs e)
        {
            string strCboValue = ((ComboBox)sender).SelectedValue.ToString();
            string strCboName = ((ComboBox)sender).Name;
            switch (strCboName)
            {
                case "cboName":
                    this.txtName.Text = strCboValue;
                    break;
                case "cboIdCard":
                    this.txtIdCard.Text = strCboValue;
                    break;
                case "cboArea":
                    this.txtArea.Text = strCboValue;
                    break;
                case "cboBirthday":
                    this.txtBirthday.Text = strCboValue;
                    break;
                case "cboSex":
                    this.txtSex.Text = strCboValue;
                    break;
            }
            
        }
        /// <summary>
        /// 读取配置文件
        /// </summary>
        private void LoadSeting()
        {

            this.txtName.Text = iniInfo.ReadString("RegString", "IdNameRegex", "", configIniPath);
            this.txtIdCard.Text = iniInfo.ReadString("RegString", "IdCardRegex", "", configIniPath);
            this.txtArea.Text = iniInfo.ReadString("RegString", "AreaRegex", "", configIniPath);
            this.txtBirthday.Text = iniInfo.ReadString("RegString", "BrithRegex", "", configIniPath);
            this.txtSex.Text = iniInfo.ReadString("RegString", "SexRegex", "", configIniPath);
            this.rbtnRange.Checked = Convert.ToBoolean(iniInfo.ReadString("PicSize", "rbtnRange", "", configIniPath));
            this.rbtnComment.Checked = Convert.ToBoolean(iniInfo.ReadString("PicSize", "rbtnComment", "", configIniPath));
            this.cboWidth.Text = iniInfo.ReadString("PicSize", "RangeColWidth", "", configIniPath);
            this.cboHeight.Text = iniInfo.ReadString("PicSize", "RangeRowHeight", "", configIniPath);
        }
        /// <summary>
        /// 保存配置文件
        /// </summary>
        private void  SaveSeting()
        {
            iniInfo.WriteString("RegString", "IdNameRegex", txtName.Text, configIniPath);
            iniInfo.WriteString("RegString", "IdCardRegex", txtIdCard.Text, configIniPath);
            iniInfo.WriteString("RegString", "AreaRegex", txtArea.Text, configIniPath);
            iniInfo.WriteString("RegString", "BrithRegex", txtBirthday.Text, configIniPath);
            iniInfo.WriteString("RegString", "SexRegex", txtSex.Text, configIniPath);

            iniInfo.WriteString("PicSize", "RangeColWidth", this.cboWidth.Text, configIniPath);
            iniInfo.WriteString("PicSize", "RangeRowHeight", this.cboHeight.Text, configIniPath);
            iniInfo.WriteString("PicSize", "rbtnRange",Convert.ToString(this.rbtnRange.Checked), configIniPath);
            iniInfo.WriteString("PicSize", "rbtnComment", Convert.ToString(this.rbtnComment.Checked), configIniPath);

        }


        private void btnSetRegex_Click(object sender, EventArgs e)
        {
            SaveSeting();
            MessageBox.Show("参数设置已更改！");
        }

  




        
    }
}
