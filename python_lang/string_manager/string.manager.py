import string
import random

str_list = [
'user1 - 		id1',
'user2 - 	id2',
'user3 - 	id3'
]

letters = string.ascii_letters
print ( ''.join(random.choice(letters) for i in range(10)) )

for s in str_list:
    substr = s.split('-')
    s1 = substr[0].lstrip()
    s2 = substr[0].lstrip()
    password = ''.join(random.choice(letters) for i in range(6))
    print(f'                      ;{s1}          ; group1,group2  ; {password}; {s2}')