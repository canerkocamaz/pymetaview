import xml.dom.minidom
import zipfile
file=zipfile.ZipFile("deneme.docx","r")
z=file.read("docProps/app.xml")
dom=xml.dom.minidom.parseString(z)
print "AppVersion :",dom.getElementsByTagName("AppVersion")[0].childNodes[0].data
print "Temlate :",dom.getElementsByTagName("Template")[0].childNodes[0].data
print "TotalTime :",dom.getElementsByTagName("TotalTime")[0].childNodes[0].data
print "Pages :",dom.getElementsByTagName("Pages")[0].childNodes[0].data
print "Words :",dom.getElementsByTagName("Words")[0].childNodes[0].data
print "Pages :",dom.getElementsByTagName("Pages")[0].childNodes[0].data
print "Characters :",dom.getElementsByTagName("Characters")[0].childNodes[0].data
print "Application :",dom.getElementsByTagName("Application")[0].childNodes[0].data
print "DocSecurity:",dom.getElementsByTagName("DocSecurity")[0].childNodes[0].data
print "Lines :",dom.getElementsByTagName("Lines")[0].childNodes[0].data
print "Paragraphs :",dom.getElementsByTagName("Paragraphs")[0].childNodes[0].data
print "ScaleCrop :",dom.getElementsByTagName("ScaleCrop")[0].childNodes[0].data
print "Company :",dom.getElementsByTagName("Company")[0].childNodes[0].data
print "LinksUpToDate :",dom.getElementsByTagName("LinksUpToDate")[0].childNodes[0].data
print "CharactersWithSpaces :",dom.getElementsByTagName("CharactersWithSpaces")[0].childNodes[0].data
print "SharedDoc :",dom.getElementsByTagName("SharedDoc")[0].childNodes[0].data
print "HyperlinksChanged :",dom.getElementsByTagName("HyperlinksChanged")[0].childNodes[0].data
print "AppVersion :",dom.getElementsByTagName("AppVersion")[0].childNodes[0].data



