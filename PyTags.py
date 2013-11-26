#!/usr/bin/python
#
# MsTag.py :: Python Library for Microsoft Tags API.
# This library is designed to interact with the Microsoft Tags API.
# Specs on https://ws.tag.microsoft.com/MIBPService.wsdl
# It's based on the Lightweight SOAP client suds.
#
# Useful Links -> http://www.jansipke.nl/python-soap-client-with-suds
# Antonio Davoli - davo 2013
# http://davo.li

import logging
import base64
from datetime import datetime

try:
    import suds
except ImportError:
    print "Sorry, you don't have the suds client installed, and this"
    print "script relies on it. Please install or reconfigure the suds"
    print "client (http://pypi.python.org/pypi/suds) and try again."


# Tag Class
class Tag:
    statesRange = ['Active', 'Paused']
    typesRange = ['pdf', 'wmf', 'jpeg', 'png', 'gif', 'tiff', 'tag']

    def __init__(self, title, url, status='Active', startDate=None):
        assert(title), "Tag title must be not Null"
        assert(url), "Tag title must be not Null"
        self.interactionNote = None
        self.types = None
        self.title = title
        self.medFiUrl = url
        self.UTCStartDate = None
        self.UTCEndDate = None
        if (status in self.statesRange):
            self.status = status
        self.UTCStartDate = None

# MsTag Class
class PyTag:

    wsdl = "https://ws.tag.microsoft.com/MIBPService.wsdl"
    options = {'Category': 'Category', 'ImageType': 'ImageType',
        'Tag': 'Tag', 'URITag': 'URITag', 'FreeTextTag': 'FreeTextTag',
        'VCardTag': 'VCardTag', 'DialerTag': 'DialerTag',
        'UserCredential': 'UserCredential'}
    decorationTypes = ['HCCBRP_DECORATION_NONE', 'HCCBRP_DECORATION_DOWNLOAD',
        'HCCBENCODEFLAG_STYLIZED', 'HCCBRP_DECORATION_FRAMEPLAIN',
        'HCCBRP_DECORATION_TEXT']
    imageTypes = ['pdf', 'wmf', 'jpeg', 'png', 'gif', 'tiff', 'tag']
    access_token = None
    client = None
    MSTagException = None
        
    def __init__(self, at):
        assert (at), "Access Token must be not Null"    
        self.client = suds.client.Client(self.wsdl)            
        self.access_token = at
        
    def create_category(self, category):
        uc = self.client.factory.create('ns1:UserCredential')
        uc.AccessToken = self.access_token
        c = self.client.factory.create('ns1:Category')
        c.Name = category
        c.UTCStartDate = datetime.now().isoformat()
        c.CategoryStatus = 'Active'
        try:
            response = self.client.service.CreateCategory(uc, c)
        except suds.WebFault, e:
            print "[MsTag] An error occured during the Category creation"
            print e
    
    def activate_category(self, category):
        uc = self.client.factory.create('ns1:UserCredential')
        uc.AccessToken = self.access_token
        try:
            response = self.client.service.ActivateCategory(uc, category)
        except suds.WebFault, e:
            print "[MsTag] An error occured during the Category Activation"
            print e
    
    def update_category(self, old_category, new_category):
        uc = self.client.factory.create('ns1:UserCredential')
        uc.AccessToken = self.access_token
        c = self.client.factory.create('ns1:Category')
        c.Name = new_category
        c.UTCStartDate = datetime.now().isoformat()
        c.CategoryStatus = 'Active'
        try:
            response = self.client.service.UpdateCategory(uc, old_category, c)
        except suds.WebFault, e:
            print "[MsTag] An error occured during the Category update"
            print e    

    def pause_category(self, category):
        uc = self.client.factory.create('ns1:UserCredential')
        uc.AccessToken = self.access_token
        try:
            response = self.client.service.PauseCategory(uc, category)
        except suds.WebFault, e:
            print "[MsTag] An error occured during the Category Pause operation"
            print e

    def create_tag(self, categoryName, tag):
        uc = self.client.factory.create('ns1:UserCredential')
        uc.AccessToken = self.access_token
        t = self.client.factory.create('ns1:URITag')
        t.Title = tag.title
        t.MedFiUrl = tag.medFiUrl
        t.UTCStartDate = datetime.now().isoformat()
        t.Status = 'Active'
        try:
            response = self.client.service.CreateTag(uc, categoryName, t)
        except suds.WebFault, e:
            print "[MsTag] An error occured during the Tag creation"
            print e            
    

    def delete_tag(self, categoryName, tag_name):
        uc = self.client.factory.create('ns1:UserCredential')
        uc.AccessToken = self.access_token
        try:
            response = self.client.service.DeleteTag(uc, categoryName, tag_name)
        except suds.WebFault, e:
            print "[MsTag] An error occured during the Tag Deletion"
            print e
    

    def activate_tag(self, categoryName, tagName):
        uc = self.client.factory.create('ns1:UserCredential')
        uc.AccessToken = self.access_token
        try:
            response = self.client.service.ActivateTag(uc, categoryName, tagName)
        except suds.WebFault, e:
            print "[MsTag] An error occured during the Tag Activation"
            print e
    
    def pause_tag(self, categoryName, tagName):
        uc = self.client.factory.create('ns1:UserCredential')
        uc.AccessToken = self.access_token
        try:
            response = self.client.service.PauseTag(uc, categoryName, tagName)
        except suds.WebFault, e:
            print "[MsTag] An error occured during the Tag Activation"
            print e
    
        
    def get_barcode(self, categoryName, tag, barcode_fname, image_type, size_in_inches, 
                    decorationsType='HCCBRP_DECORATION_NONE', is_bw=False):
        if ((image_type in self.imageTypes) == False):
            print "[MsTag] Image Type not allowed."
            return
        if ((decorationsType in self.decorationTypes) == False):
            print "[MsTag] Decoration Type not allowed"
            return
        
        uc = self.client.factory.create('ns1:UserCredential')
        uc.AccessToken = self.access_token
        im = self.client.factory.create('ns1:ImageTypes')
        im = image_type
        dt = self.client.factory.create('ns1:DecorationType')
        dt = decorationsType
        try:
            response = self.client.service.GetBarcode(uc, categoryName, tag, im, size_in_inches, dt, is_bw)
        except suds.WebFault, e:
            print "[MsTag] An error occured during the Tag Activation"
            print e
        try:
            o = open(barcode_fname, "wb")
        except IOError as (errno, strerror):
            print "I/O error({0}): {1}".format(errno, strerror)
            
        o.write(base64.b64decode(self.client.last_received().getChild("s:Envelope").getChild("s:Body").getChild("GetBarcodeResponse").getChild("GetBarcodeResult").getText()))
        o.close()
        return
