# convert2txt

 pdf, docx, pptxファイルからテキストを抽出し、各抽出結果をtxtファイルに エクスポートする 

## Requirement

* pdfminer.six >= 20200726
* python-pptx >= 0.6.18
* docx2txt >= 0.8

## Installation

```bash
pip install -r requirements.txt
```

 

## Usage

 `./file_list/`に対象ファイル(`.pdf`, `pptx`, `docx`)を配置します。

``` 
python convert2txt.py
```

上記コマンド実行により、`all_pdf.txt`,`all_pptx.txt`, `all_docx`が出力されます。

## Note

- 出力ファイルは実行ごとに上書き
- 1ファイルごとに1行の空行が入る
- pdfについては元ファイルの見た目ベースで改行されてしまう



 

 