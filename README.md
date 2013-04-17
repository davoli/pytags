pytags
======

Python Library for Microsoft Tags API  

How To:
------

After the library import, create a new MsTag object 
passing the token obtained by the Microsoft Tags Website.

    myAuthToken = "you-token-here"
    m = PyTag(myAuthToken)
    
Create a new Category and proceed with the activation.

    r = m.create_category("TestCategoryNew")
    r = m.activate_category("TestCategory")
    
Create a new Tag and get the Pdf file (in this case)

    tag = Tag("DavoTag", "http://davo.li")
    m.create_tag("TestCategory", tag)
    m.activate_tag('TestCategory', "DjangoTag")
    m.get_barcode('TestCategory', "DavoTag", "DavoTag.pdf", 'pdf', 2)
