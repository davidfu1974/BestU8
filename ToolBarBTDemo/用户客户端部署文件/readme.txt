-- 运行下面的数据库脚本 UFMeta_xxxx 库 (17代表销售订单)，可以通过ctl+shit 点击工具条按钮后将内容复制到写字板查看
INSERT INTO [AA_CustomerButton]([cButtonID], [cButtonKey], [cButtonType], [cProjectNO], [cFormKey], [cVoucherKey], [cKeyBefore], [iOrder], [cGroup], [cCustomerObjectName], [cCaption], [cLocaleID], [cImage], [cToolTip], [cHotKey], [bInneralCommand], [cVariant], [cVisibleAsKey], [cEnableAsKey])
VALUES(newid(), 'ToolBarBTDemo','default', 'U8CustDef','17', '17','save', '0', 'IEDIT','ToolBarBTDemo.ToolBarBTDemoClass','测试按钮','zh-cn','','测试','Ctrl+N',1,'测试按钮点击后参数数据传入','save','save')

-- 复制用户自定义开发的COM按钮组件到 C:\U8SOFT\UAP\RUNTIME  或 C:\U8SOFT\PORTAL
C:\U8SOFT\Portal\Interop.UAPVoucherControl85.dll
C:\U8SOFT\Portal\ToolBarBTDemo.dll
C:\U8SOFT\Portal\ToolBarBTDemo.tlb
--C:\U8SOFT\UAP\RUNTIME\ToolBarBTDemo.dll
--C:\U8SOFT\UAP\RUNTIME\Interop.UAPVoucherControl85.dll




-- 注册COM组件
C:\Windows\Microsoft.NET\Framework\v2.0.50727\regasm C:\U8SOFT\Portal\ToolBarBTDemo.dll /tlb C:\U8SOFT\Portal\ToolBarBTDemo.tlb /codebase
--C:\Windows\Microsoft.NET\Framework\v2.0.50727\regasm C:\U8SOFT\UAP\RUNTIME\ToolBarBTDemo.dll /tlb:C:\U8SOFT\UAP\RUNTIME\ToolBarBTDemo.tlb /codebase 
--C:\U8SOFT>gacutil -i C:\U8SOFT\UAP\RUNTIME\ToolBarBTDemo.dll
--C:\Windows\Microsoft.NET\Framework\v4.0.30319>regasm C:\U8SOFT\UAP\RUNTIME\\ToolBarBTDemo.dll


--卸载注册组件
C:\Windows\Microsoft.NET\Framework\v2.0.50727\regasm /u C:\U8SOFT\Portal\ToolBarBTDemo.dll /tlb C:\U8SOFT\Portal\ToolBarBTDemo.tlb /codebase
--C:\U8SOFT>gacutil -u ToolBarBTDemo
--C:\Windows\Microsoft.NET\Framework\v4.0.30319>regasm  /u C:\U8SOFT\UAP\RUNTIME\\ToolBarBTDemo.dll

--开发注意点：
1、项目采用C#类库，.NET 平台最好选择3.5 
2、项目属性release 为 X86平台
3、生成选项中，勾选“为COM互操作注册”，之后VS需要管理员启动编译。
4、项目属性中应用程序中，“是程序集COM可见”。
5、项目引用应该引用 C:\U8SOFT\ufcomsql\UAPvouchercontrol85.ocx
6、开发可以参考文档
7、目前看来用户自定义开发的组件只能放在portal 目录下，否则会调用出错，感觉找不到路径。