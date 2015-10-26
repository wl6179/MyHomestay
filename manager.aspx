<%@ Page language="c#" Codebehind="manager.aspx.cs" AutoEventWireup="false" Inherits="MyHomestay.manager" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" >
<HTML>
	<HEAD>
		<title>manager</title>
		<meta name="GENERATOR" Content="Microsoft Visual Studio .NET 7.1">
		<meta name="CODE_LANGUAGE" Content="C#">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
			<asp:DataGrid id="GameOneDayDataGrid" style="Z-INDEX: 101; LEFT: 8px; POSITION: absolute; TOP: 112px"
				runat="server" Width="1314px" AutoGenerateColumns="False" Font-Size="X-Small" BorderColor="Tan"
				BorderWidth="1px" BackColor="LightGoldenrodYellow" CellPadding="2" GridLines="None" ForeColor="Black"
				DataKeyField="UserID">
				<SelectedItemStyle ForeColor="GhostWhite" BackColor="DarkSlateBlue"></SelectedItemStyle>
				<AlternatingItemStyle BackColor="PaleGoldenrod"></AlternatingItemStyle>
				<HeaderStyle Font-Bold="True" BackColor="Tan"></HeaderStyle>
				<FooterStyle BackColor="Tan"></FooterStyle>
				<Columns>
					<asp:TemplateColumn HeaderText="���">
						<ItemTemplate>
							<%# GetgoCount()%>
						</ItemTemplate>
					</asp:TemplateColumn>
					<asp:BoundColumn DataField="User_RName" HeaderText="����"></asp:BoundColumn>
					<asp:BoundColumn DataField="User_Sex" HeaderText="�Ա�" ReadOnly="True"></asp:BoundColumn>
					<asp:BoundColumn DataField="User_age" HeaderText="��������" ReadOnly="True"></asp:BoundColumn>
					<asp:BoundColumn DataField="User_City" HeaderText="����"></asp:BoundColumn>
					<asp:BoundColumn DataField="User_Add" HeaderText="��ַ"></asp:BoundColumn>
					<asp:BoundColumn DataField="User_Email" HeaderText="E��Mail" ReadOnly="True"></asp:BoundColumn>
					<asp:BoundColumn DataField="User_MPhone" HeaderText="�ֻ�"></asp:BoundColumn>
					<asp:BoundColumn DataField="User_Phone" HeaderText="�绰"></asp:BoundColumn>
					<asp:BoundColumn DataField="User_Wsex" HeaderText="����Ա�" ReadOnly="True"></asp:BoundColumn>
					<asp:BoundColumn DataField="User_Chinese" HeaderText="�������" ReadOnly="True"></asp:BoundColumn>
					<asp:BoundColumn DataField="User_populace" HeaderText="�˿�" ReadOnly="True"></asp:BoundColumn>
					<asp:BoundColumn DataField="User_house" HeaderText="���ݽṹ" ReadOnly="True"></asp:BoundColumn>
					<asp:BoundColumn DataField="User_net" HeaderText="����" ReadOnly="True"></asp:BoundColumn>
					<asp:BoundColumn DataField="User_pet" HeaderText="����" ReadOnly="True"></asp:BoundColumn>
					<asp:BoundColumn DataField="User_Date" HeaderText="ע������" ReadOnly="True"></asp:BoundColumn>
					<asp:EditCommandColumn ButtonType="PushButton" UpdateText="����" HeaderText="�༭" CancelText="ȡ��" EditText="�༭"></asp:EditCommandColumn>
				</Columns>
				<PagerStyle HorizontalAlign="Center" ForeColor="DarkSlateBlue" BackColor="PaleGoldenrod"></PagerStyle>
			</asp:DataGrid>
			<asp:Label id="Label1" style="Z-INDEX: 102; LEFT: 320px; POSITION: absolute; TOP: 40px" runat="server">�û���</asp:Label>
			<asp:Label id="Label2" style="Z-INDEX: 103; LEFT: 328px; POSITION: absolute; TOP: 80px" runat="server">����</asp:Label>
			<asp:TextBox id="TextBox1" style="Z-INDEX: 104; LEFT: 400px; POSITION: absolute; TOP: 40px" runat="server"></asp:TextBox>
			<asp:TextBox id="TextBox2" style="Z-INDEX: 105; LEFT: 400px; POSITION: absolute; TOP: 72px" runat="server"
				TextMode="Password"></asp:TextBox>
			<asp:Button id="Button1" style="Z-INDEX: 106; LEFT: 584px; POSITION: absolute; TOP: 72px" runat="server"
				Width="72px" Text="��¼"></asp:Button>
			<asp:LinkButton id="GONEnextButton" style="Z-INDEX: 107; LEFT: 736px; POSITION: absolute; TOP: 72px"
				runat="server" Font-Size="X-Small" CommandName="gonext">��һҳ</asp:LinkButton>
			<asp:LinkButton id="GONEupButton" style="Z-INDEX: 108; LEFT: 680px; POSITION: absolute; TOP: 72px"
				runat="server" Font-Size="X-Small" CommandName="goup">��һҳ</asp:LinkButton>
			<asp:Label id="GONEnullLabel" style="Z-INDEX: 109; LEFT: 800px; POSITION: absolute; TOP: 72px"
				runat="server" Font-Size="X-Small"></asp:Label>
			<asp:Label id="GONEshowLabel" style="Z-INDEX: 110; LEFT: 904px; POSITION: absolute; TOP: 72px"
				runat="server" Font-Size="X-Small"></asp:Label>
		</form>
	</body>
</HTML>
