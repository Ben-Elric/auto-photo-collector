# auto-photo-collector

## 介绍
一个结合腾讯文档收集图片并打包的简单工具，主要用于完成上级学院/团委/学校组织的行政任务。<br/>
主要应用场景：收集青年大学习等截图。<br/>
本项目结合QQ小程序腾讯文档中收集表收集截图功能，对导出工作表的截图url的图片爬取下载并用表格对应名字命名，而后收集打包，并给出未打卡名单。<br/>

## 部署说明
### 环境
pip install requests和zipfile库
### 部署
1.收集姓名并修改list.txt为下图样式<br/>
2.使用腾讯文档收集表收集截图，并导出excel工作表。<br/>
3.将工作表命名（或者无需）成带有“大学习”字段的文件，放入脚本同一目录。(如要更换命名需求，自行修改代码)。<br/>
4.双击收集大学习.py 。

### 结果
具体见[pdf](Readme.pdf)。

## 注意事项!
1.跑脚本时需要关掉梯子。<br/>
2.文件均存放在英文目录。

## 附录
感谢使用。<br/>
关注我的博客：[iben7.xyz](https://iben7.xyz/)。



