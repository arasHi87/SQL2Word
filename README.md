# SQL2Word

Convert SQL table to Word table.

## Installation

```bash
git clone https://github.com/arasHi87/SQL2Word
pip3 install -r requirements.txt 
```

## Usage

Run `main.py` with following args, and it will auto save in pwd with name `meow.docx`

* host: your database host/ip
* user: your database username
* passwd: your database password
* db_name: your database name
* ssl_ca_path: the path where your ssl ca
* ssl_cert_path: the path where your ssl cert
* ssl_key_path: the path where your ssl key

```bash
python3 main.py <host> <user> <passwd> <db_name> <ssl_ca_path> <ssl_cert_path> <ssl_key_path>
```
