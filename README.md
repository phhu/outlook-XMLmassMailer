This is an extension for Microsoft Outlook which allows mass mailing to be done based on data in an XML file.

![Mass mailer form](/img/massMailerForm.png)

Installation
============

* Open the visual basic editor (Outlook, Alt-F11)
* Import the .bas and .frm files

![Visual basic modules](/img/modules.png)

![Import file](/img/importFile.png)

* Add references to 
** "Microsoft XML, v6.0"
** "Microsoft Office x.0 Object library"
** "Microsoft Word x.0 Object library"

![References](/img/references.png)

* Save
* Then run the macro "makeToolbar"

Use
===

XML files look like something like this...

```XML
<root>
	<row>
		<to>somebody@company.com</to>
		<cc>somebody@company.com</cc>
		<bcc>somebody@company.com</bcc>
		<subject>somebody@company.com</subject>
		<body>Some text here</body>
	</row>
	<row>
		<to>somebody@company.com</to>
		<cc>somebody@company.com</cc>
		<bcc>somebody@company.com</bcc>
		<subject>somebody@company.com</subject>
		<htmlbody><p>Some text here</p></htmlbody>
	</row>	
</root>
```