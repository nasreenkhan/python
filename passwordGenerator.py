import random

lower = "abcdefghijklmnopqrstuvwxyz"
upper = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
numbers = "0123456789"
symbols = "[]{}()*;/-_@#$"

mixed = lower + upper + numbers + symbols

length = 16

password = "".join(random.sample(mixed,length))

print("password: ", password)