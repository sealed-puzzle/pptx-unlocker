import re
import sys
import zipfile

FIND_RE = re.compile(r'<p:modify.*?/>')

def unlock_pptx(src_file, dest_file=None):
  if dest_file is None:
    dest_file = re.sub(r'(\.[^.]+)$', r'_patch\1', src_file)

  target_file = 'ppt/presentation.xml'

  with zipfile.ZipFile(src_file, 'r') as zin:
    with zipfile.ZipFile(dest_file, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
      zout.comment = zin.comment
      for x in zin.infolist():
        if x.filename != target_file:
          zout.writestr(x, zin.read(x.filename))

      with zin.open(target_file, 'r') as fin:
        with zout.open(target_file, 'w') as fout:
          for line in fin:
            line = re.sub(FIND_RE, '', line.decode('utf-8')).encode('utf-8')
            fout.write(line)

if __name__ == '__main__':
  if len(sys.argv) != 2:
    print('Usage: python3 unlock_pptx.py <file>', file=sys.stderr)
  else:
    unlock_pptx(sys.argv[1])
