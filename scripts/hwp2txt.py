"""HWP 5.x 문서의 본문 텍스트를 출력한다.  사용법: python3 hwp2txt.py 파일.hwp"""
import olefile, zlib, struct, sys

CTRL16 = (1, 2, 3, 4, 5, 6, 7, 8, 11, 12, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23)

def hwp_text(path):
    ole = olefile.OleFileIO(path)
    comp = ole.openstream('FileHeader').read()[36] & 1
    secs = sorted([s for s in ole.listdir() if s[0] == 'BodyText'], key=lambda s: int(s[1][7:]))
    out = []
    for s in secs:
        d = ole.openstream(s).read()
        d = zlib.decompress(d, -15) if comp else d
        i = 0
        while i < len(d):
            h = struct.unpack_from('<I', d, i)[0]; tag = h & 0x3ff; size = (h >> 20) & 0xfff; i += 4
            if size == 0xfff:
                size = struct.unpack_from('<I', d, i)[0]; i += 4
            r = d[i:i + size]; i += size
            if tag != 67:   # HWPTAG_PARA_TEXT
                continue
            t, j = '', 0
            while j < len(r):
                c = struct.unpack_from('<H', r, j)[0]
                if c in CTRL16: j += 16; continue
                if c < 32: j += 2; continue
                t += chr(c); j += 2
            out.append(t)
    return '\n'.join(out)

if __name__ == '__main__':
    print(hwp_text(sys.argv[1]))
