import json
file_path = "orgDistinctMeetings.txt"
with open(file_path) as f:
    for line in f:
        j_content = json.loads(line)
        print(j_content)
