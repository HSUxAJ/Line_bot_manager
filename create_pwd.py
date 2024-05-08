import random
import string

def generate_random_string(length):
    # 使用 string 模块中的 digits、ascii_lowercase 和 ascii_uppercase 来获取数字和字母
    characters = string.digits + string.ascii_lowercase + string.ascii_uppercase
    # 使用 random 模块的 sample 函数从 characters 中随机选择 length 个字符
    random_string = ''.join(random.sample(characters, length))
    return random_string

# 生成一个包含所有数字和英文字母的 20 位随机字符串
random_string = generate_random_string(23)
print(random_string)
