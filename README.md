# **This is a plug-in based of easyui.** #
This is a plug-in based of easyui, This plug-in can export datagrid to Excel, The effect is what you see is what you get, not only contains data but also contains styles, contains the style of easyui itself and user defined.  
---
这是一个基于easyui的插件。这个插件可以将datagrid导出成excel，效果是所见即所得，不仅仅只是数据，还包括样式，包含easyui自己的样式和用户自定义的部分。  
---
include:
```
<script type="text/javascript" src="../easyui-datagrid-export.js"></script>
```
---
use demo:
```
$("#datagrid_table_id").datagrid("toExcel", "filename.xls");
```