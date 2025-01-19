def find_command_and_return_index(line: str, command: str):
    words = list(line.split())
    idx = words.index(command)
    if command == 'ㅁㅁ':
        stdnamelist = list()
        score = int()
        detail = str()
        for i in range(idx+1, len(words)):
            # print(words[i])
            if ('-' in words[i]) or ('+' in words[i]):
                score = int(words[i].replace('점', ''))
                detail = words[i+1]
                break
            else:
                stdnamelist.append(words[i])
        return [stdnamelist, score, detail]
    else:
        stdname = words[idx+1]
        score = int(words[idx+2].replace('점', ''))
        detail = words[idx+3]
        return [stdname, score, detail]