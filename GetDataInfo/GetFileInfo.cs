using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows.Forms;
 
namespace GetDataInfo
{
    class GetFileInfo
    {

   
        /// <summary>
        /// 弹出对话框，选择多个文件。
        /// 为true时获取选择文件的详细路径，为false时只获取文件的名称和扩展名
        /// </summary>
        /// <param name="bolFullPath">为true时获取选择文件的详细路径，为false时只获取文件的名称和扩展名</param>
        /// <returns>返回文件的详细路径或只返回文件的名称和扩展名</returns>
        public List<string> GetFilePath(bool bolFullPath)
        {

            List<string> dirFilePath = new List<string>();
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "图片 (*.BMP;*.JPG;*.JPEG;*.PNG;*.GIF)|*.BMP;*.JPG;*.JPEG;*.PNG;*.GIF|" + "所有文件 (*.*)|*.*";
            ofd.RestoreDirectory = true;//指示该对话框在关闭前是否将目录还原为之前选定的目录
            ofd.Multiselect = true;
            ofd.FileName = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            //ofd.ShowDialog();
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                if (bolFullPath)
                {
                  
                    foreach (var item in ofd.FileNames)
                    {
                       
                        dirFilePath.Add(item);
                    }                                       
                }
                else
                {
                    foreach (var item in ofd.SafeFileNames)
                    {
                        dirFilePath.Add(item);
                    }     
                }


                return dirFilePath;
            }
            ofd.Dispose();
            return null;           
        }
        /// <summary>
        /// 提取一个完整路径的文件中的指定类型信息
        /// </summary>
        /// <param name="filePathName">完整的文件路径</param>
        /// <param name="style">类型：0-获取文件完整路径，1-获取文件扩展名，2-获取文件名没有扩展名也没有路径，3-只得到文件名并包含扩展名，4-只得到文件所在目录下的路径 </param>
        /// <returns></returns>
        public string GetFileStyle(string filePathName,int style)
        {
            switch (style)
            {
                 case 0:
                     return System.IO.Path.GetFullPath(filePathName); //文件完整路径
                 case 1:
                    return System.IO.Path.GetExtension(filePathName);  //文件扩展名
                 case 2:
                    return System.IO.Path.GetFileNameWithoutExtension(filePathName);//文件名没有扩展名也没有路径
                 case 3:
                   return  System.IO.Path.GetFileName(filePathName);//只得到文件名并包含扩展名
                 case 4:
                    return   System.IO.Path.GetDirectoryName(filePathName);//只得到文件所在目录下的路径  
            }

            return null;
        }

        /// <summary>
        /// 弹出对话框，选择文件夹路径
        /// </summary>
        /// <returns>返回所选文件夹的所在路径</returns>
        public string GetFolderPath()
        {
            string folderPath;
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.Description = "请选择文件夹";
            //dialog.ShowDialog();
            if (dialog.ShowDialog() == DialogResult.OK)
            { 
                folderPath = dialog.SelectedPath;
                return folderPath;
            }
            dialog.Dispose();
            return null;
            
        }



    }
}
