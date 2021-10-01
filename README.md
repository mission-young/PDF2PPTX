从PPTX转PDF的方式很多，通过ppt自带的工具即可实现快速而精确的转换。但从PDF转PPTX的转换却缺乏简单便捷的方式。

该问题的需求在于：

  - ppt编辑公式比较繁琐，且需要花费大量的时间用来调节字体等格式，且不美观

  - latex编辑公式较为方便，但生成的PDF与课堂白板sewoo并不兼容，无法很好地利用画笔在课件上进行标记，给授课带来不遍。

一个简单的想法是看看有无现有的工具将PDF直接导出为PPTX，但发现无论是adobe acrobat还是foxit pdf reader，亦或者PDF24 tools均喜欢做额外的工作，进行OCR识别，效果还不甚理想。直接转化为图片再拼接成PPTX不香么！多此一举！

好在PDF24 tools可以将PDF拆分成图片，而PPT又可以很方便地导入图片。

![image](https://upload-images.jianshu.io/upload_images/12062705-32f7954284c14e0a.png?imageMogr2/auto-orient/strip%7CimageView2/2/w/1240)

![image](https://upload-images.jianshu.io/upload_images/12062705-1f1bee2858f4ee60.png?imageMogr2/auto-orient/strip%7CimageView2/2/w/1240)

这种方式可以有效地解决我的需求。但我想更简单点，有没有一键完成的可能？

最初的想法是使用VBA，随后导出为exe可执行程序。但缺乏现有的教程指导。随后查询Github，找到了几个pdf转png的项目，结合网上的部分资源，耗时一个晚上，实现了一键tex->pdf->pngs->pptx的工作流。


``` python

# python

# windows

# pyinstaller -i pdf2pptx.ico --add-data=".\default.pptx;." -F -w pdf2pptx.py

# macOS

# pyinstaller -i pdf2pptx.ico --add-data=".\default.pptx:." -F  pdf2pptx.py

# pipenv 下打包缩小体积

# pip install pyinstaller pymupdf python-pptx PyPDF2

import fitz # pip install  pymupdf

import os

import sys

import shutil

import pptx #pip install python-pptx

import PyPDF2

#生成资源文件目录访问路径

def resource_path(relative_path):

    if getattr(sys, 'frozen', False): #是否Bundle Resource

        base_path = sys._MEIPASS

    else:

        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

def pdf2pptx(pdf_name,zoom):

    # pdf2pngs

    pdf_name=os.path.abspath(pdf_name)

    if pdf_name[-4:].lower() == '.pdf':

        dir_name = os.path.dirname(pdf_name) # 获得地址的父链接

        base_name = os.path.basename(pdf_name)[0:-4] # 获得地址的文件名

        path = dir_name + os.sep + base_name

        if not os.path.exists(path):

            os.makedirs(path)

        with open(pdf_name,'rb') as f:

            pdfwidth=PyPDF2.PdfFileReader(f).getPage(0).mediaBox[2]

            pdfheight=PyPDF2.PdfFileReader(f).getPage(0).mediaBox[3]

        pdfscale=pdfheight/pdfwidth

        pdfmode = 0

        if pdfscale > 0.74 and pdfscale < 0.76:

            pdfmode = 0

        elif pdfscale > 0.5 and pdfscale < 0.6:

            pdfmode = 1

        else:

            pdfmode = -1

        pdf = fitz.open(pdf_name)

        # pdfwidth=pdf[0].get_images()[0][2]

        # pdfheight=pdf[0].get_images()[0][3]

        # pdfscale=pdfheight/pdfwidth

        # pdfmode = 0

        # if pdfscale > 0.74 and pdfscale < 0.76:

        #    pdfmode = 0

        # elif pdfscale > 0.5 and pdfscale < 0.6:

        #    pdfmode = 1

        # else:

        #    pdfmode = -1

        for pg in range(0, pdf.pageCount):

            page = pdf[pg]  # 获得每一页的对象

            trans = fitz.Matrix(zoom, zoom).preRotate(0)

            pm = page.getPixmap(matrix=trans, alpha=False)  # 获得每一页的流对象

            pm.writePNG(path + os.sep + base_name + '_' + '{:0>3d}.{}'.format(pg+1, 'png'))  # 保存图片

        pdf.close()

        print('PDF转PNG成功!')

    # pngs2pptx

        pngs=os.listdir(path+os.sep)

        template=resource_path(os.path.join("default.pptx"))

        prs=pptx.Presentation(template)

        if pdfmode == 0:

            prs.slide_width = pptx.util.Inches(4)

            prs.slide_height = pptx.util.Inches(3)

        elif pdfmode == 1:

            prs.slide_width = pptx.util.Inches(16)

            prs.slide_height = pptx.util.Inches(9)

        else:

            print('请采用标准格式!')

            prs.slide_width = pptx.util.Inches(4)

            prs.slide_height = pptx.util.Inches(3)

        layout=prs.slide_layouts[6]

        for png in pngs:

            slide=prs.slides.add_slide(layout)

            slide.shapes.add_picture(path + os.sep + png,0,0,height=prs.slide_height)

        prs.save(path + '.pptx')

        print('PNG转PPTX成功!')

        shutil.rmtree(path)

    else:

        print('警告! 文件类型错误,请打开pdf文件类型!')

if __name__=="__main__":

    if len(sys.argv) == 2:

        pdf2pptx(sys.argv[1],5)

    elif len(sys.argv) > 2:

        pdf2pptx(sys.argv[1],int(sys.argv[2]))

    else:

        print('请输入要转化的文件名')

```

在此过程中，又研究了如何将python源文件编译为独立的exe文件。

中途遇到了几个问题，列如下：

  - 相对路径和绝对路径的问题。为了避免前后不一致，在程序运行前段将文件地址设定为绝对路径。

  - 获取PDF文件的长宽比。PPT文件有16:9和4:3两种比例，通过或许原PDF的长宽比，可以设定PPTX的长宽。最初的方案直接利用fitz库即可提取，但部分PDF会发生错误。因而改用现有PyPDF2库来提取该特征。此外，最初的想法是将长宽比作为参数输入，但后来觉得太蠢，还是通过上述方法来实现，更能体现自动化。

  - 将python源文件编译为exe后，找不到模板文件"default.pptx"。这一点经常出现在pyinstaller编译中，由于脱离了原python运行环境，部分资源无法找到，会出现这种错误。因而需要将"default.pptx"文件拷贝到工作目录，并通过`pptx.Presentation("default.pptx")`的方式来调用。但这种方式会导致需要将该模板与exe文件绑定，既不方便，也不优雅。随后查阅pyinstaller手册，可以通过`--add-data`参数，将该文件内嵌到exe文件内部。

  - 程序运行期间，会出现cmd窗口。 编译命令中添加`-w`参数解决。

  - 编译生成的exe文件较大。可以考虑使用pipenv环境进行编译，可以降低空间占用。

  - 环境配置。安装fitz库，直接安装pymupdf包即可。安装pptx库，直接安装python-pptx包。

完成以上步骤之后，将pdf文件用该可执行文件打开，便可以转化为pptx文件。

为了进一步自动化，每次在tex编译生成pdf之后，自动将pdf转为pptx，查阅了vscode插件`Latex Workshop`的手册。

![image](https://upload-images.jianshu.io/upload_images/12062705-dd49670cdf7ff7ad.png?imageMogr2/auto-orient/strip%7CimageView2/2/w/1240)

注意到适当调整占位符，即可完成自动将pdf转为pptx的过程。

随后修改配置文件：

```json

"latex-workshop.latex.tools": [

      {

              "name": "pdf2pptx",

              "command": "pdf2pptx",

              "args": [

              "%DOCFILE%.pdf"

              ]

      }

],

"latex-workshop.latex.recipes": [

      {

              "name": "xelatex",

              "tools": [

              "xelatex",

              "pdf2pptx"

          ]

      }

] 

```

其中`latex-workshop.latex.tools`对应`latex-workshop.latex.recipes`中的`tools`键.

至此，编辑tex文件之后，只需保存一下，即可自动生成pdf和pptx文件。

程序：[pdf2pptx.exe](https://github.com/mission-young/PDF2PPTX/releases/download/v1/pdf2pptx.exe)

Github项目地址：[mission-young/PDF2PPTX (github.com)](https://github.com/mission-young/PDF2PPTX)
