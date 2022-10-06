f = open("student_id.txt")

id = []

for i in range(1, 79):
    s = f.readline()
    s = s[:-1]
    id.append(s)

f = open("student_name.txt")
name = []
for i in range(1, 79):
    s = f.readline()
    s = s[:-1]
    name.append(s)

mp = {}
for i in range(0, 78):
    mp[name[i]] = id[i]