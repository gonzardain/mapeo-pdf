import os
base = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
fonts = os.path.join(base, "fonts/")

def funcion1():
	print fonts + 'AvenirNextLTPro-Bold.otf'
	funcion2()

def funcion2():
	print base
	funcion3()

def funcion3():
	print base


if __name__ == "__main__":
	funcion1()