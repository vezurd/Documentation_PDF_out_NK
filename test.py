def get_paege_count(page_format):
    f_letters = ["A","А"]
    s = page_format.replace(" ", "")
    s = s.split(",")
    count = 0
    for i in range(len(s)):
        print(s[i])
        if s[i][0] in f_letters:
            count += 1
        else:
            end_index = 0
            for a in range(len(f_letters)):
                index = s[i].find(f_letters[a])
                if index > 0:
                    end_index = int(index)
            print (end_index)

            count = count + int(s[i][0:end_index])
    return count
print(get_paege_count("А0, А3х3, А4х3, А4х3"))