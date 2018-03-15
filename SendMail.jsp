<%@page import="java.io.File"%>
<%@page import="org.eclipse.swt.SWT"%>
<%@page import="org.eclipse.swt.ole.win32.OLE"%>
<%@page import="org.eclipse.swt.ole.win32.OleAutomation"%>
<%@page import="org.eclipse.swt.ole.win32.OleClientSite"%>
<%@page import="org.eclipse.swt.ole.win32.OleFrame"%>
<%@page import="org.eclipse.swt.ole.win32.Variant"%>
<%@page import="org.eclipse.swt.widgets.Display"%>
<%@page import="org.eclipse.swt.widgets.Shell"%>
<%@ page language="java" contentType="text/html; charset=ISO-8859-1"
    pageEncoding="ISO-8859-1"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<title>Insert title here</title>
</head>
<body>
<script type="text/javascript">
function sendPDF() {
<%
    	Display display = Display.getDefault();
    	Shell shell = new Shell(display);
    	OleFrame frame = new OleFrame(shell, SWT.NONE);
    	OleClientSite site2 = new OleClientSite(frame, SWT.NONE,
    	        "Outlook.Application");
    	OleAutomation outlook = new OleAutomation(site2);
    	OleAutomation mail = outlook.invoke(outlook.getIDsOfNames(new String[] { "CreateItem" })[0],new Variant[] { new Variant(0) }) .getAutomation();
    	mail.setProperty(mail.getIDsOfNames(new String[] { "To" })[0], new Variant("test@gmail.com"));
    	mail.setProperty(mail.getIDsOfNames(new String[] { "Bcc" })[0], new Variant("test@gmail.com"));
    	mail.setProperty(mail.getIDsOfNames(new String[] { "Subject" })[0], new Variant("Top News for you"));
    	mail.setProperty(mail.getIDsOfNames(new String[] { "HtmlBody" })[0], new Variant("<html>Hello<p>, please find some infos here.</html>"));
    	mail.setProperty(mail.getIDsOfNames(new String[] { "BodyFormat" })[0], new Variant(2));
    	File file = new File("C:/Users/" + System.getProperty("user.name") + "/Downloads/SalPal_.pdf");
    	OleAutomation attachments = null;
    	if (file.exists()) {
    		Variant varResult = mail.getProperty(mail.getIDsOfNames(new String[] { "Attachments" })[0]);
    	    if (varResult != null && varResult.getType() != OLE.VT_EMPTY) {
    	    	attachments = varResult.getAutomation();
    	        varResult.dispose();
    	    }
    	    attachments.invoke(attachments.getIDsOfNames(new String[] { "Add" })[0],new Variant[] { new Variant("C:/Users/" + System.getProperty("user.name") + "/Downloads/SalPal_.pdf") });
    	} else {
    	}
    	mail.invoke(mail.getIDsOfNames(new String[] { "Display" })[0]);

    	if(!display.isDisposed())
        {
        	display.close();
        }
%>
}
</script>
</body>
</html>