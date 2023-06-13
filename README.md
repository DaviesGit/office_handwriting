[![MIT License][license-shield]][license-url]



<br />
<p align="center">
  <a href="https://github.com/DaviesGit">
    <img src="readme_images/Ideal_Logo_Davies.ico" alt="Logo" width="150">
  </a>

  <h3 align="center">Maked by Davies</h3>

  <p align="center">
    Email: 1182176003@qq.com
<!--     <br />
    <a href="https://github.com/DaviesGit"><strong>Explore the docs »</strong></a>
    <br />
    <br />
    <a href="javascript:void(0)">View Demo</a>
    ·
    <a href="javascript:void(0)">Report Bug</a>
    ·
    <a href="javascript:void(0)">Request Feature</a> -->
  </p>
</p>



<!-- TABLE OF CONTENTS -->
## Table of Contents

* [About the Project](#about-the-project)
  * [Built With](#built-with)
* [Getting Started](#getting-started)
  * [Prerequisites](#prerequisites)
  * [Installation](#installation)
* [Usage](#usage)
* [功能定制](#功能定制)
* [Roadmap](#roadmap)
* [Contributing](#contributing)
* [License](#license)
* [Contact](#contact)
* [Acknowledgements](#acknowledgements)
* [免责声明](#免责声明)


<!-- ABOUT THE PROJECT -->
## About The Project

打印效果

![小红帽](test/小红帽.jpg)



源文档 [小红帽](https://daviesgit.github.io/office_handwriting/test/小红帽.pdf)
<p align="center">
    <iframe width="100%" height="800px" src="test/小红帽.pdf"></iframe>
</p>



手写文章生成脚本，可模仿手写字体。



功能:

* 模仿手写字体



### Built With
依赖
* [Microsoft Office](https://www.office.com/)



<!-- GETTING STARTED -->

## Getting Started

这个章节将指导你简单的部署和使用该软件。

### Prerequisites

这个项目的依赖安装步骤在下面给出。
* [Microsoft Office](https://www.office.com/)

> 请下载最新版[Microsoft Office](https://www.office.com/)



### Installation

** 注意：中文版 office 请使用 export_cn 文件夹中的代码！！！ **

1. Clone the repo
```sh
git clone https://github.com/path/to/the/repository
```

2. 安装`handwriting_font_config`文件夹内的所有字体。
3. 打开`word`选择`视图`>`宏`>`查看宏`

![result00](readme_images/result00.png)

4. 填写宏的名称`handwriting`点击`创建`

   ![result01](readme_images/result01.png)


5. 将`handwriting.vba`文件中的内容复制到vbs编辑器中保存

   ![result02](readme_images/result02.png)


6. 右键`Class Modules`选择`插入`>`Class Module`

   ![result03](readme_images/result03.png)


7. 将`FontConfig.vba`中的内容复制到编辑器中保存，并将类名改为`FontConfig`

   ![result04](readme_images/result04.png)





<!-- USAGE EXAMPLES -->

## Usage

1. 打开`查看宏`的窗口，选择刚刚创建的宏，点击执行。

   ![result05](readme_images/result05.png)

2. 手动调整部分字体格式。

3. 进行你需要的操作。

4. Good luck!




## 功能定制

如果需要功能定制，请联系作者 [1182176003@qq.com](1182176003@qq.com)




<!-- ROADMAP -->
## Roadmap

See the [open issues](https://example.com) for a list of proposed features (and known issues).



<!-- CONTRIBUTING -->
## Contributing

Contributions are what make the open source community such an amazing place to be learn, inspire, and create. Any contributions you make are **greatly appreciated**.

1. Fork the Project
2. Create your Feature Branch (`git checkout -b feature/AmazingFeature`)
3. Commit your Changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the Branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request



<!-- LICENSE -->
## License

Distributed under the MIT License. See `LICENSE` for more information.



<!-- CONTACT -->
## Contact

Davies - [@qq](1182176003) - 1182176003

Davies - [@email](1182176003@qq.com) - 1182176003@qq.com

Project Link: [https://example.com](https://example.com)



<!-- ACKNOWLEDGEMENTS -->
## Acknowledgements
* [GitHub](https://github.com/)
* [Font Awesome](https://fontawesome.com)



## 免责声明
* 该软件中所包含的部分内容，包括文字、图片、音频、视频、软件、代码、以及网页版式设计等可能来源于网上搜集。

* 该软件提供的内容仅用于个人学习、研究或欣赏，不可使用于商业和其它意图，一切关于该软件的不正当使用行为均与我们无关，亦不承担任何法律责任。使用该软件应遵守相关法律的规定，通过使用该软件随之而来的风险与我们无关，若使用不当，后果均由个人承担。

* 该软件不提供任何形式的保证。我们不保证内容的正确性与完整性。所有与使用该软件的直接风险均由用户承担。

* 如果您认为该软件中所包含的部分内容侵犯了您的权益，请及时通知我们，我们将尽快予以修正或删除。


<!-- MARKDOWN LINKS & IMAGES -->
<!-- https://www.markdownguide.org/basic-syntax/#reference-style-links -->

[license-shield]: readme_images/MIT_license.svg
[license-url]: https://opensource.org/licenses/MIT

[product-screenshot]: readme_images/screenshot.png
