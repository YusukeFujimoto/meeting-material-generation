"""
PowerPoint COM (win32com) 経由でスライドをPNGに書き出し、ビジュアルQAに使うためのスクリプト。

このWindows環境ではLibreOffice (soffice) が入っておらず、pptxスキル標準の
scripts/office/soffice.py は `socket.AF_UNIX` が無いというエラーで必ず失敗する
（AF_UNIXソケットを使うサンドボックス回避ロジックがLinux/macOS前提のため）。
一方、この機体にはMicrosoft Officeが入っているので、pywin32
（`import win32com.client` が通ることを事前に確認しておく。通らなければ
`pip install pywin32`）でPowerPoint本体を自動操作し、実際のフォント・レイアウトで
書き出す方が確実。LibreOfficeでのレンダリングはYu Mincho/Yu Gothicのようなユーザー環境の
日本語フォントを代替フォントに置き換えてしまい、QAで見ている見た目が実際の成果物と
ズレることがある——このスクリプトなら実機PowerPointと同じ見え方を確認できる。

使い方:
    python export_slides_win32.py <deck.pptx> <output_dir>

出力: <output_dir>/slide-1.png, slide-2.png, ...
"""
import os
import sys

sys.stdout.reconfigure(encoding="utf-8")
import win32com.client as win32


def export(path, outdir, width=1600, height=900):
    path = os.path.abspath(path)
    outdir = os.path.abspath(outdir)
    os.makedirs(outdir, exist_ok=True)

    ppt = win32.gencache.EnsureDispatch("PowerPoint.Application")
    try:
        pres = ppt.Presentations.Open(path, WithWindow=False)
        outputs = []
        for i, slide in enumerate(pres.Slides, start=1):
            outpath = os.path.join(outdir, f"slide-{i}.png")
            slide.Export(outpath, "PNG", width, height)
            outputs.append(outpath)
            print("exported", outpath)
        pres.Close()
        return outputs
    finally:
        ppt.Quit()


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("usage: python export_slides_win32.py <deck.pptx> <output_dir>")
        sys.exit(1)
    export(sys.argv[1], sys.argv[2])
