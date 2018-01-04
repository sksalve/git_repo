import os

from xml.dom.minidom import parseString, parse


class XML:
	path =r'/home/administrator/Desktop/xmlwatcher'
	xmls= {}

	def get_xml(xmldoc, faxid, email, stamptime, pages, status):
		for xml in os.listdir(XML.path):
			if xml.endswith('.xml'):
				xmldoc = parse(xml)
				faxid = xmldoc.getElementsByTagName('FaxID')[0].firstChild.data
	            email = xmldoc.getElementsByTagName('BillingCode')[0].firstChild.data
	            stamptime = xmldoc.getElementsByTagName('CustomCode2')[0].firstChild.data
	            pages = xmldoc.getElementsByTagName('Pages')[0].firstChild.data
	            status = xmldoc.getElementsByTagName('Status')[0].firstChild.data
	      	return (faxid, email, stamptime, pages, status)

