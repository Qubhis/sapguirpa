# SapGuiEpa

I'm making this repository public as I think that it might help some people. I personally use this for developing my own scripts for various cases such as:
 - process testing
 - master data maintenance where BAPI or LSMW cannot help
 - new plant code creation with additional setting, especially EC02 technical log scrapping
 - transactional data creation for new functionality testing (if not using GUI's built-in recording)
 - prototyping of RPA robots

Please note that the documentation of the code is not the best. The code can be optimized and some parts can be changed for better readability, but before I get to that (and if), I will appreciate any comments. Questions are welcomed as well :). 

### Quick usage from a command line:
```python
>>> from sapguirpa import SapGuiRpa
>>> sap = SapGuiRpa()
>>> sap.attach_to_session()
```
![image of select session gui popup]()
```python
>>> sap.start_transaction("ME21N")
```
