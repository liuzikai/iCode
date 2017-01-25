你好 <(￣ˇ￣)/ 欢迎使用iCode

iCode目前尚处测试状态，可能会出现各种Bug，使用有一定风险。

如有任何问题或建议，非常欢迎与我联系，共同努力改善它。

为了避免损失，请注意工程文件的保存与备份。使用本软件造成的损失开发者不负任何责任。

开发者：liuzikai
邮箱：liuzikai@163.com
贴吧：蓝极光弧


◇iCode的安装与卸载

iCode自动安装至VB6安装目录下“iCode”文件夹内，可从“程序与功能”中卸载iCode


◇iCode的开启与关闭

VB-外接程序-外接程序管理器-iCode，“加载/卸载”选项即是开关，勾选“在启动中加载”则会随VB一起启动。


◇备份VB界面布局设置

接入iCode可能导致VB界面上部分按钮在下一次启动“消失”，原因尚不明确。为了避免重新布局的麻烦，可提前备份VB界面布局设置。

贴吧链接：http://tieba.baidu.com/p/3964446594?pid=73519240602&cid=0#73519240602
1.打开注册表编辑器（运行regedit）
2.定位至HKEY_CURRENT_USER\Software\Microsoft\VBA\Microsoft Visual Basic，在左侧栏中的“Microsoft Visual Basic”上单右键，选择导出所选分支为reg文件。
3.定位至HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0，同样导出为reg文件。
如果出现VB界面布局异常，导入这两个reg文件即可恢复。


◇减小风险
1.注意工程文件的保存与备份。
2.通过“选项”关闭工作区标签栏（出现Bug的可能性比较高，其次是文件窗口重布局）。

