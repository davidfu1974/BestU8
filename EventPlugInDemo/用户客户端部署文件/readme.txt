1.打开API资源管理器
2.U8API ->销售管理->销售订单->事件->插件类型->保存前事件，注册同步插件。
3. 注册同步事件：
插件编码：EventSaleOrderSaveBeforePlugInDemo
插件名称：EventSaleOrderSaveBeforePlugInDemo
插件绑定：勾选修改后选择DONETASSEMBLYFORRPC 
指定如下信息
AssemblyPath:  C:\templogs\EventPlugInDemo.dll
ClassFullName:  EventPlugInDemo.EventSaleOrderSaveBeforePlugInDemo
MethodName:  Save_before

4、确定后自动匹配成功即可

