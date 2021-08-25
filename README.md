### 项目简介 
本项目脚本是处理企业工商银行的电子对账单，将其转化为会计方便使用的格式。

### 使用方法
- 新建虚拟环境
```bash
virtualenv -p python3 venv
source bin/activate
```
将项目解压到venv下
修改 ORIGINAL_FILE, OUTPUT_FILE

- 初始化
```bash
cd seven701
sudo python3 -m pip install -r requirements.txt
python3 changecsv2xlsx.py
```


mine todo
导出的原始文件，直接放过来执行，看看能否执行成功？
如果成功，那就再把csv文件处理，放到git上；
如果不行，再改代码
