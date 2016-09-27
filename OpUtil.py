from OpRemain_process import *
from Op2Serve_process import *

if __name__ == '__main__':
	if len(sys.argv) > 1 :
		os.chdir(os.path.dirname(sys.argv[1]))
		table = xlReader(*sys.argv)()
		hdr, *tup = table
		inspect_idx = hdr.index('불출일자')
		if tup[0][inspect_idx] == '':
			op = Op2Serve(table)()
			Op2ServeSaveTo(op)()
		else:
			op=OpRemain(table)()
			OpRemainSaveTo(op)()
			


