# auto-extract-data
* Clone repo này: ```git clone https://github.com/quangnm145/auto-extract-data```
* Cài đặt python: 3.13.4

* Chạy lệnh sau để cài đặt tất cả package từ file: ```pip install -r requirements.txt ```

* Cách dùng:
```
usage: docx_to_md.py [-h] [-i INPUT] [-c PASSWORD] [-o OUTPUT]

Chuyển file .docx (có hoặc không có mật khẩu) sang .md, bao gồm cả bảng

options:
  -h, --help            show this help message and exit
  -i, --input INPUT     File .docx đầu vào
  -c, --password PASSWORD
                        Mật khẩu của file .docx (nếu có)
  -o, --output OUTPUT   File .md đầu ra (mặc định: tên file đầu vào với đuôi .md)

Ví dụ: python docx_to_md.py -i test.docx -c 1a@ -o test.md 
       python docx_to_md.py -i test.docx 
       python docx_to_md.py -h
```
