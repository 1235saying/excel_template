# yaml 配置文件可能用的更加广泛


import yaml

with open("hello.yaml", "r") as f:
    data = yaml.safe_load(f)

print(type(data))

print(data)