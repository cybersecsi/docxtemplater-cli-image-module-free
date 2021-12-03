docxtemplater CLI (with Image Module Free)
==========================================
This fork of docxtemplater CLI includes the support for [docxtemplater-image-module-free](https://github.com/evilc0des/docxtemplater-image-module-free). This module is different from the module [open-docxtemplater-image-module](https://github.com/MaxRcd/open-docxtemplater-image-module), which does not work now.
Moreover, this cli also includes the support for the rendering of some basic html code. This code is taken from [pwndoc](https://github.com/pwndoc/pwndoc) and the syntax to use it is the same used by pwndoc itself and is documented [here](https://pwndoc.github.io/pwndoc/#/docxtemplate?id=html-values-from-text-editors).

Installation
------------
The package below includes a copy of docxtemplater and docxtemplater-image-module-free.
```
npm install -g docxtemplater-cli-image-module-free
```

Try it out
----------
Browse to folder 'examples' and run the following in your console
```
docxtemplater input.pptx data.json output.pptx
```

data.json file structure
------------------------
This is a simple JSON file. Assuming we are working on the template 'input.pptx' provided with this package in 'examples' folder, we write below JSON to replace tags in the template by their respective value.
```json
{
    "first_name": "John",
    "last_name": "Doe",
    "description": "He is awesome",
    "phone": "+4412345678",
    "picture": "johndoe.png"
}
```

As you can see we have included a picture of John Doe (johndoe.png). For this picture to be correctly rendered in the output, we must explicitly activate the Image Module Free [docxtemplater-image-module-free](https://github.com/evilc0des/docxtemplater-image-module-free). Note that all images must be under the directory 'imageDir'.
```json
{
    "first_name": "John",
    "last_name": "Doe",
    "description": "He is awesome",
    "phone": "+4412345678",
    "picture": "johndoe.png",
    "config": {
      "modules": ["docxtemplater-image-module-free"],
      "imageDir": "."
    }
}
```

Template tags
-------------
More on how to write your templates in [docxtemplater documentation](http://docxtemplater.readthedocs.io/en/latest/tag_types.html).