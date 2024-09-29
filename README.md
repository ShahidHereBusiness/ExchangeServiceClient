# ExchangeServiceClient
Read and Send Email from microsoftonline.com

# EWS
EWService.asmx.cs
Line 17: Set Namespace="https://domain/SOA/"

web.config
Add Inside Tag <configuration></configuration>

  <appSettings>
    <add key="uri" value="https://domain/ews/exchange.asmx"/>
    <add key="EWSSpecifier" value="emailid@domain"/>
    <add key="EWSWords" value="watchWord"/>
    <add key="path" value="DriveLetter:\\DevLOGS\\"/>
    <add key="accessSpecifier" value="role"/>
    <add key="watchWord" value="pass"/>
    <add key="encryptItem" value="True"/>
    <add key="AESKey" value="ed1b64d08c9da16a80d3971340c7d698"/>
    <add key="AESiv" value="d68c69039117Ma10"/>
    <add key="WhiteListed" value="::1|127.0.0.1"/>
    <add key="Token" value="d9de4a84dd6b80a1971c09d163c78306"/>
  </appSettings>

Service Definition:
EWService.asmx?WSDL

Service URI:
EWService.asmx

Header:
Content-Type=text/xml

Request Body:
<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
  <soap:Body>
    <GetItem xmlns="https://domain/SOA/">
      <accessSpecifier>role</accessSpecifier>
      <watchWord>pass</watchWord>
      <paramsList>TO|FROM|SUBJECT|BODY|DTR</paramsList>
      <item>
	   <Flag>0</Flag>
		<Id></Id>
		<From></From>
		<To></To>
		<Headers></Headers>
		<Subject></Subject>
		<Body></Body>
		<DTR></DTR>
      </item>
    </GetItem>
  </soap:Body>
</soap:Envelope>

Response:
<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">
    <soap:Body>
        <GetItemResponse xmlns="https://domain/SOA/">
            <GetItemResult>0</GetItemResult>
            <item>
                <Flag>0</Flag>
                <Id />
                <From />
                <To />
                <Headers />
                <Subject />
                <Body />
                <DTR />
            </item>
        </GetItemResponse>
    </soap:Body>
</soap:Envelope>
