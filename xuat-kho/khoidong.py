"""Start dependencies, server and browser through one verified Python runtime."""
import os
from pathlib import Path
import subprocess
import sys
import traceback

ROOT = Path(__file__).resolve().parent

def main():
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, 'reconfigure'):
            stream.reconfigure(encoding='utf-8', errors='replace')
    os.chdir(ROOT)
    try:
        if sys.version_info < (3, 10):
            raise RuntimeError('Cần Python 3.10 trở lên. Hãy cài Python mới rồi chạy lại.')
        print('Đang kiểm tra Python và thư viện Excel...', flush=True)
        print(f'Python {sys.version.split()[0]}', flush=True)
        try:
            import openpyxl
            parts=tuple(int(v) for v in openpyxl.__version__.split('.')[:3])
            ready=(3,1,5)<=parts<(4,0,0)
        except ImportError:
            ready=False
        if not ready:
            subprocess.run([sys.executable,'-m','pip','install','-r',str(ROOT/'requirements.txt')],check=True)
        from app import serve
        # Try a nearby port if another program is using 8080.
        for port in range(8080,8090):
            try:
                serve(port=port,open_browser=True)
                return 0
            except OSError as exc:
                if getattr(exc,'winerror',None)==10048 or exc.errno==98:
                    continue
                raise
        raise RuntimeError('Các cổng 8080–8089 đều đang được sử dụng. Hãy đóng phiên ứng dụng cũ và thử lại.')
    except KeyboardInterrupt:
        return 0
    except Exception:
        details=traceback.format_exc()
        try:
            (ROOT/'Loi-khoi-dong.log').write_text(details,encoding='utf-8')
        except OSError:
            pass
        print('\nKhông khởi chạy được ứng dụng:\n'+details,flush=True)
        print('Gửi nội dung trên hoặc file Loi-khoi-dong.log để kiểm tra.',flush=True)
        return 1

if __name__=='__main__':
    raise SystemExit(main())
