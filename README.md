# 从msi安装包文件中获取属性数据实现安装包的重命名

写这篇文章的起因是，女朋友在实现CI时，最后一步是生成一个安装包归档。生成安装包使用的工具是Visual Studio Installer，这个工具在VS2017之后已经不再默认提供，当然还是可以以VS插件的形式加载进来使用，可以通过“工具-扩展和更新”在线下载或者[离线下载](https://marketplace.visualstudio.com/items?itemName=VisualStudioClient.MicrosoftVisualStudio2017InstallerProjects)再安装进来。使用这个工具制作的安装包，可以指定安装包名称和版本号，但是有一个蛋疼的地方在于，生成的安装包的名字没有版本宏可以附加，导致的结果就是生成的安装包的名字没法带版本号。人工去改安装包名加上版本号就体现不出CI的优势了。她和我分享了这么一个故事，她的解决思路是用python之类的脚本文件解析生成安装包的工程文件.vdproj文件，然后找出版本对应的字段，在生成安装包后用这个字段改名。我从另一个角度去想了这个问题，于是，就有了这篇文章。

## 如何使用

假设制作的安装包中，给定的安装包名称为InstallerDemo，版本号为2.0.0.100，用Visual Studio Installer生成的安装包为InstallerDemo.msi。
- 执行命令`FormatMsi.exe InstallerDemo.msi`后该安装包将被重命名为InstallerDemo_2.0.0.100.msi。
- 或直接运行FormatMsi.exe，然后输入安装包名InstallerDemo.msi，之后该安装包也会被重命名为InstallerDemo_2.0.0.100.msi。


代码在GitHub上的[FormatMsi](https://github.com/KondeU/FormatMsi)中开源，有相应的[Release包](https://github.com/KondeU/FormatMsi/releases)可以下载，也可以参考本文。代码依赖.Net Framework 4.5.2。

## 关键源码

代码很简单，C#写的，结合注释看看就懂了。核心部分在GetPropertyFromMsi函数中，通过WindowsInstaller解析msi文件，读出属性数据，再根据属性改名。

如果有自定义的需求，可以参考Main中formatMsiFileName的上下文修改。

```cpp
////////////////////////////////////////////////////////////////////////////////
//
// MIT License
//
// Copyright (c) 2021 kongdeyou(https://tis.ac.cn/blog/kongdeyou/)
//
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in all
// copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
// SOFTWARE.
//
////////////////////////////////////////////////////////////////////////////////

using System;
using System.IO;
using WindowsInstaller;

namespace FormatMsi
{
    class Program
    {
        static void Main(string[] args)
        {
            String inputFile = null;

            if (args.Length == 1)
            {
                inputFile = args[0];
            }
            else
            {
                Console.WriteLine("Please enter the msi file:");
                inputFile = Console.ReadLine();
            }

            String productName = null;
            String productVersion = null;
            String formatMsiFileName = null;
            try
            {
                if (inputFile.EndsWith(".msi", StringComparison.OrdinalIgnoreCase))
                {
                    productName = GetPropertyFromMsi(inputFile, "ProductName");
                    productVersion = GetPropertyFromMsi(inputFile, "ProductVersion");
                }
                else
                {
                    Console.WriteLine("Error: Invalid input file!");
                    return;
                }

                formatMsiFileName = String.Format("{0}_{1}.msi", productName, productVersion);
                Console.WriteLine("Format msi file name: " + formatMsiFileName);

                File.Copy(inputFile, formatMsiFileName);
                File.Delete(inputFile);
            }
            catch (Exception exception)
            {
                Console.WriteLine("Exception: " + exception.Message);
            }
        }

        static String GetPropertyFromMsi(String msi, String property)
        {
            String ret = null;

            // WindowsInstaller from [SYSTEM]:\Windows\System32\msi.dll
            Type classType = Type.GetTypeFromProgID("WindowsInstaller.Installer");
            Installer installer = Activator.CreateInstance(classType) as Installer;

            // Open the msi file for reading, 0 means read, 1 means read and write
            Database database = installer.OpenDatabase(msi, 0);

            // The requested property fetching command
            String sql = String.Format(
                "SELECT Value FROM Property WHERE Property='{0}'", property);

            // Open the database view and then execute SQL command
            View view = database.OpenView(sql);
            view.Execute(null);

            // Read from the fetched record
            Record record = view.Fetch();
            if (record != null)
            {
                ret = record.get_StringData(1);
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(record);
            }

            // Close the database view
            view.Close();

            // Release the view's and the database's COM object
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(view);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(database);

            return ret;
        }
    }
}
```

## 添加引用

除了以上的代码外，我们在代码中使用了一个关于msi的库，即WindowsInstaller，它在C#项目模板中默认是不带的，我们需要把它加进来。

在解决方案资源管理器中找到对应的项目，这里就是FormatMsi，然后在“引用”上右键，选择“添加引用”，然后会弹出一个对话框，选择最下方的“浏览”，找到系统目录下的`Windows\System32\msi.dll`文件，添加进来。

## 协议

本文以上内容遵循CC BY-ND 4.0协议，署名-禁止演绎。

本文中的源代码遵循MIT开源协议。
代码托管于：<https://github.com/KondeU/FormatMsi>

转载请注明出处：<https://tis.ac.cn/blog/kongdeyou/format_msi/>

作者：[kongdeyou(https://tis.ac.cn/blog/author/kongdeyou/)](https://tis.ac.cn/blog/author/kongdeyou/)
