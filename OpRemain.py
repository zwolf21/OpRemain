from OpRemain_process import *



if __name__ == '__main__':
	if len(sys.argv) > 1 :
		os.chdir(os.path.dirname(sys.argv[1]))
		table = xlReader(*sys.argv)()
		op=OpRemain(table)()
		OpRemainSaveTo(op)()
