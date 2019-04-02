import os

def iter_files(path):
	"""Walk through all files located under a root path."""
	if os.path.isfile(path):
		yield path
	elif os.path.isdir(path):
		for dirpath, _, filenames in os.walk(path):
			for f in filenames:
				yield os.path.join(dirpath, f)
	else:
		raise RuntimeError('Path %s is invalid' % path)


path = "pdf path"
exe_path='pdf2html.exe path'
output_path = 'output path'

if not os.path.exists(output_path):
	os.makedirs(output_path)
bats = ''
for file in iter_files(path):
	cmd = "%s -i %s -o %s\n" % (exe_path,file,output_path)
	bats+=cmd
bats+='echo \"process all sucessfully\"\n'
with open('run.bat','w') as f:
	f.write(bats)