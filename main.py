#from functions import GetRightNumber, GetRightNumber2 , GetRightNumber3
from functions import GetRightNumber2 , GetRightNumber3
from app import Application

#funcs = [GetRightNumber(),GetRightNumber2(),GetRightNumber3()]
funcs = [GetRightNumber2(),GetRightNumber3()]
application = Application(funcs)
application.mainloop()