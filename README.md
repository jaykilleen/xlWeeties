# xlWeeties

Simply copy and paste the code from [init.bas](https://github.com/jaykilleen/xlWeeties/blob/master/init.bas) into a new module named 'xlWeeties'.

Also copy the [add_reference.bas](https://github.com/jaykilleen/xlWeeties/blob/master/add_references.bas) to a new module called `add_references`.

For it to all work you will need to add the `Microsoft Visual Basic For Applications Extensibility 5.3` library to your VBA Editor (as `add references` cannot do this for you because it depends on this library).

You can do this as per [cpearson](http://www.cpearson.com/excel/vbe.aspx) instructions:

>First, you need to set an reference to the VBA Extensibility library. The library contains the definitions of the objects that make up >the VBProject. In the VBA editor, go the the Tools menu and choose References. In that dialog, scroll down to and check the entry for >Microsoft Visual Basic For Applications Extensibility 5.3. If you do not set this reference, you will receive a User-defined type not >defined compiler error.

You can then run the `add_references` module.

You can then run the `xlWeeties` module.

:)
