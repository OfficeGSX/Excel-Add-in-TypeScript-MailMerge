# <a name="excel-add-in-typescript-mailmerge"></a>Excel 外接程序 TypeScript 邮件合并

用于 TypeScript 的 Excel 邮件合并外接程序连接到 Microsoft Graph，从 Outlook 中的模板文件夹中获取电子邮件模板，并发送来自 Excel 表中的收件人列表的邮件。

![开始页](../readme-images/first_run.PNG)

## <a name="prerequisites"></a>先决条件

若要运行示例，将需要：

* Visual Studio 2015
* Microsoft Visual Studio TypeScript 最低版本 2.0.6.0 
* [Node.js](https://nodejs.org/)
* Office 365 开发人员帐户。如果你没有帐户，可以[加入 Office 365 开发人员计划并获取为期 1 年的免费 Office 365 订阅](https://aka.ms/devprogramsignup)。

## <a name="run-the-add-in"></a>运行外接程序

### <a name="register-your-app-in-microsoft-azure"></a>在 Microsoft Azure 中注册应用

使用以下配置在[应用注册门户](https://apps.dev.microsoft.com)中注册 Web 应用程序：

参数 | 值
---------|--------
名称 | Excel-Add-in-Microsoft-Graph-MailMerge
类型 | Web 应用程序和/或 Wb API
登录 URL | https://localhost:44390/index.html
应用 ID URI | https://[你的 Azure AD 租户名].onmicrosoft.com/Excel-Add-in-Microsoft-Graph-MailMerge
回复 URL | https://localhost:44390/index.html

添加以下权限：

应用程序 | 委派权限
---------|--------
Microsoft Graph | 读/写邮件
Microsoft Azure Active Directory | 登录并读取用户个人资料

保存应用程序并记下*客户端 ID*。

### <a name="set-up-your-environment"></a>设置环境

1. 克隆 GitHub 存储库。
3. 在 Visual Studio 中，打开解决方案文件 Excel-Add-in-Microsoft-Graph-MailMerge.sln。

### <a name="update-the-client-id"></a>更新客户端 ID

* 在你的 Visual Studio 项目中，打开 Excel-Add-in-Microsoft-Graph-MailMergeWeb/src/home/home.ts。
* 使用你 Azure AD 应用程序中的值更新“[在此处输入你的客户端 ID]”。
* 使用你的重定向 URL 更新“[重定向 URL]”。

### <a name="run-the-add-in"></a>运行外接程序

1. 打开\<示例目录\> \Excel-Add-in-Microsoft-Graph-MailMergeWeb 的命令提示符并运行 `npm install`，完成之后运行 `npm start`。
2. 在 Visual Studio 中，按 F5 运行示例。
3. Excel 打开时，从“主页”选项卡中选择“**邮件合并**”命令按钮。

![命令按钮](../readme-images/command_button.PNG)

4. 任务窗格将会打开，单击“**登录 Microsoft**”后即可使用 Office 365 凭据进行身份验证。
5. 从模板列表中选择。

![选择模板](../readme-images/select_template.PNG)

6. 查看并编辑收件人列表。

![编辑收件人](../readme-images/mailmerge_table.PNG)

7. 预览和发送电子邮件。

![预览和发送电子邮件](../readme-images/preview_send.PNG)

## <a name="questions-and-comments"></a>问题和意见

我们乐意倾听你对此示例的反馈。可以在该存储库中的“[问题](https://github.com/OfficeDev/Excel-Add-in-TypeScript-MailMerge/issues)”部分将问题和建议发送给我们。

与 Office 365 开发相关的问题一般应发布到 [Stack Overflow](http://stackoverflow.com/questions/tagged/office-addins)。确保你的问题或意见使用 [Office 外接程序] 进行了标记。

## <a name="additional-resources"></a>其他资源

* [Office 外接程序示例](https://github.com/OfficeDev?utf8=%E2%9C%93&query=-add-in)
* [Office 外接程序平台概述](http://dev.office.com/docs/add-ins/overview/office-add-ins)
* [开始使用 Office 外接程序](http://dev.office.com/getting-started/addins)
* [Office JavaScript API 帮助程序](https://github.com/OfficeDev/office-js-helpers)

## <a name="copyright"></a>版权

版权所有 © 2016 Microsoft Corporation。保留所有权利。





