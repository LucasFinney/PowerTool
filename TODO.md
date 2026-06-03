TODO

Fix Error: 
Exception in Tkinter callback
Traceback (most recent call last):
  File "C:\Users\lucfi\miniconda3\Lib\tkinter\__init__.py", line 2079, in __call__
    return self.func(*args)
           ~~~~~~~~~^^^^^^^
  File "C:\Users\lucfi\miniconda3\Lib\tkinter\__init__.py", line 862, in callit
    func(*args)
    ~~~~^^^^^^^
  File "c:\Users\lucfi\SlideTool\PowerTool\slidetool\gui.py", line 133, in <lambda>
    self.root.after(0, lambda: self._on_done_error(e))
                                                   ^
NameError: cannot access free variable 'e' where it is not associated with a value in enclosing scope