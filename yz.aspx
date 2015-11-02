<%@ Page Language="C#" AutoEventWireup="true" CodeFile="yz.aspx.cs" Inherits="yz" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
       设置需要保存为密钥的图片：<asp:FileUpload ID="FileUpload1" runat="server" />
        <br />
        <asp:Button ID="Button2" runat="server" Text="确定" onclick="Button2_Click" />
        &nbsp;&nbsp;
        <asp:Button ID="Button1" runat="server" onclick="Button1_Click" Text="图片加密" />
&nbsp;&nbsp;&nbsp;
        <asp:Button ID="Button4" runat="server" onclick="Button4_Click" Text="下载密钥" />
        <br />
        
        <br />
        校验签名是否为真，请上传需要校验的图片密钥：<asp:FileUpload ID="FileUpload2" runat="server" />
        <br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
        <asp:Button ID="Button3" runat="server" onclick="Button3_Click" Text="进行对比" />
        <br />
        校验结果：<asp:Label ID="lb_jy" runat="server" Text="Label"></asp:Label>
        <br />
        <br />
        <asp:HiddenField ID="hdfield" runat="server" />
        <br />
        word文档安全性校验：&nbsp; 
       
            <input id="File1" type="file" runat="server" />

        <br />
        <br />
        <asp:Button ID="Button5" runat="server" onclick="btnUpload_Click" Text="确定上传" />
&nbsp;&nbsp;
        <asp:Button ID="Button6" runat="server" Text="校验" onclick="Button6_Click" />
        <br />
    
    </div>
    </form>
</body>
</html>
