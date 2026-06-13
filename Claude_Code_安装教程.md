### 一、如何成功安装claude code 
claude code 本身是个命令行工具（CLI）,在终端等黑窗口输入命令指挥其工作

##### 1、安装Node.js  --------- 一个运行环境，cloude code 需要靠它才能跑起来
```安装地址
https://nodejs.org/zh-cn/download
```
安装成功后，win+r输入cmd   node -v  出现版本号，说明安装成功
![alt text](image.png)


###### 2、安装claude code 
需要使用windows powershell  右键以管理员身份运行
或者直接用cmd 输入命令
```
 输入命令 npm install -g @anthropic-ai/claude-code
 ```
![alt text](image-1.png)
##### 2、安装Git -------   版本控制工具
本人已经安装过了，跳过此步骤，不安装会报错

![alt text](image-2.png)
出现上面截图，说明已经可以运行claude code,但是需要一些验证，需修改配置文件claude.json,增加"hasCompletedOnboarding": true
![alt text](image-3.png)
![alt text](image-4.png)
![alt text](image-5.png)
成功进入，提示是否在当前目录，然后回车，会让我们去登录claude账号，因为claude是国外的大模型，一方面需要付钱买token，另外账号经常被封。我们接入国产的DeepSeek
![alt text](image-6.png)

##### 3、cc-Switch  ------  一键轻松切换大模型API
安装CC-Switch，在Git-Hub找到安装包https://github.com/farion1231/cc-switch

![alt text](image-7.png)
![alt text](image-8.png)

打开DeepSeek的API开放平台，创建API keys  复制密钥
![alt text](image-9.png)
![alt text](image-10.png)

进入CC-Switch,添加模型，选择DeepSeek，粘贴密钥
![alt text](image-11.png)


然后重新进入cmd，运行claude,不在提示登录
![](image-12.png)

![alt text](image-13.png)
![alt text](image-14.png)】

需要充钱，deepseek需要实名认证，充钱后就可以啦
![alt text](image-15.png)
##### 4、Vscode  -------  微软出的免费编辑器，长得像记事本但功能超级强大

如何在VScode中运行claude code? 为什么选择在VScode中运行，可以可视化操作

1、安装claude code 插件
![alt text](image-16.png)

2、终端执行报错、原因问deepseek

问题原因分析
- 直接原因：当你在 PowerShell 中执行 claude 命令时，系统实际尝试运行一个名为 claude.ps1 的 PowerShell 脚本。由于系统的执行策略(Execution Policy)限制，这个运行行为被阻止了。

- 安全背景：这是 Windows 的一项安全功能，就像为系统设置了一道安检。默认的 Restricted 策略禁止所有脚本运行，以防止不慎执行恶意代码。

方案一：为当前用户永久修改执行策略 (推荐)
这个方案对所有 PowerShell 窗口永久生效，但仅作用于你当前的 Windows 用户账户，不影响系统或其他用户，是开发环境的最佳实践。

以普通身份打开 PowerShell：在 Windows 开始菜单搜索 "PowerShell"，点击运行即可，无需使用管理员模式。

执行修改命令：在打开的 PowerShell 窗口中，复制并粘贴以下命令，然后按回车：
```
powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```
RemoteSigned 策略允许运行本地创建的脚本，而从网络下载的脚本则必须有可信签名才能运行。这既解决了问题，也保留了基本的安全性。

确认更改：系统会询问是否要更改执行策略，输入 Y (Yes) 并按回车确认。

验证并重新运行：操作完成后，关闭当前 PowerShell 窗口，在 VS Code 中重新打开终端，再次尝试运行 claude 命令。


![alt text](image-17.png)






