-- ������������ݿ�ű� UFMeta_xxxx �� (17�������۶���)������ͨ��ctl+shit �����������ť�����ݸ��Ƶ�д�ְ�鿴
INSERT INTO [AA_CustomerButton]([cButtonID], [cButtonKey], [cButtonType], [cProjectNO], [cFormKey], [cVoucherKey], [cKeyBefore], [iOrder], [cGroup], [cCustomerObjectName], [cCaption], [cLocaleID], [cImage], [cToolTip], [cHotKey], [bInneralCommand], [cVariant], [cVisibleAsKey], [cEnableAsKey])
VALUES(newid(), 'ToolBarBTDemo','default', 'U8CustDef','17', '17','save', '0', 'IEDIT','ToolBarBTDemo.ToolBarBTDemoClass','���԰�ť','zh-cn','','����','Ctrl+N',1,'���԰�ť�����������ݴ���','save','save')

-- �����û��Զ��忪����COM��ť����� C:\U8SOFT\UAP\RUNTIME  �� C:\U8SOFT\PORTAL
C:\U8SOFT\Portal\Interop.UAPVoucherControl85.dll
C:\U8SOFT\Portal\ToolBarBTDemo.dll
C:\U8SOFT\Portal\ToolBarBTDemo.tlb
--C:\U8SOFT\UAP\RUNTIME\ToolBarBTDemo.dll
--C:\U8SOFT\UAP\RUNTIME\Interop.UAPVoucherControl85.dll




-- ע��COM���
C:\Windows\Microsoft.NET\Framework\v2.0.50727\regasm C:\U8SOFT\Portal\ToolBarBTDemo.dll /tlb C:\U8SOFT\Portal\ToolBarBTDemo.tlb /codebase
--C:\Windows\Microsoft.NET\Framework\v2.0.50727\regasm C:\U8SOFT\UAP\RUNTIME\ToolBarBTDemo.dll /tlb:C:\U8SOFT\UAP\RUNTIME\ToolBarBTDemo.tlb /codebase 
--C:\U8SOFT>gacutil -i C:\U8SOFT\UAP\RUNTIME\ToolBarBTDemo.dll
--C:\Windows\Microsoft.NET\Framework\v4.0.30319>regasm C:\U8SOFT\UAP\RUNTIME\\ToolBarBTDemo.dll


--ж��ע�����
C:\Windows\Microsoft.NET\Framework\v2.0.50727\regasm /u C:\U8SOFT\Portal\ToolBarBTDemo.dll /tlb C:\U8SOFT\Portal\ToolBarBTDemo.tlb /codebase
--C:\U8SOFT>gacutil -u ToolBarBTDemo
--C:\Windows\Microsoft.NET\Framework\v4.0.30319>regasm  /u C:\U8SOFT\UAP\RUNTIME\\ToolBarBTDemo.dll

--����ע��㣺
1����Ŀ����C#��⣬.NET ƽ̨���ѡ��3.5 
2����Ŀ����release Ϊ X86ƽ̨
3������ѡ���У���ѡ��ΪCOM������ע�ᡱ��֮��VS��Ҫ����Ա�������롣
4����Ŀ������Ӧ�ó����У����ǳ���COM�ɼ�����
5����Ŀ����Ӧ������ C:\U8SOFT\ufcomsql\UAPvouchercontrol85.ocx
6���������Բο��ĵ�
7��Ŀǰ�����û��Զ��忪�������ֻ�ܷ���portal Ŀ¼�£��������ó����о��Ҳ���·����