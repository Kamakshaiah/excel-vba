def HelloWorldPython( ):

    desktop = XSCRIPTCONTEXT.getDesktop()

    model = desktop.getCurrentComponent()

    if not hasattr(model, “Text”):

    model = desktop.loadComponentFromURL(

    “private:factory/swriter”,”_blank”, 0, () )

    text = model.Text

    tRange = text.End

    tRange.String = “Hello World (in Python)”

    return None
