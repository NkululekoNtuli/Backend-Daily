#TODO: Create a letter using starting_letter.docx 
#for each name in invited_names.txt
#Replace the [name] placeholder with the actual name.
#Save the letters in the folder "ReadyToSend".
    
#Hint1: This method will help you: https://www.w3schools.com/python/ref_file_readlines.asp
    #Hint2: This method will also help you: https://www.w3schools.com/python/ref_string_replace.asp
        #Hint3: THis method will help you: https://www.w3schools.com/python/ref_string_strip.asp
def get_letter_file():
    with open('mail merge/input/letters/starting_letter.docx', 'r', errors= 'ignore') as file:
        letter = file.read()
    return letter

def get_names_file():
    with open('mail merge/input/Names/inivited_names.txt', 'r', errors='ignore') as file:
        invited_names = file.readlines()
        return invited_names

def get_each_name():
    invited_names = get_names_file()
    list_of_names = []
    for i in invited_names:
        if "\n" in i:
            i = i[:len(i)-1] + "," + i[len(i)-1:]
            list_of_names.append(i)
        else:
            list_of_names.append(i)
    return list_of_names
print(get_each_name())

def make_letters():
    list_of_names = get_each_name()
    list_of_letters = []

    for i in range(len(list_of_names)):
        comma = ","
        letter_template = get_letter_file()
        if "[name]," in letter_template:
            letter_template = letter_template.replace("[name],\n", list_of_names[i])
            list_of_letters.append(letter_template)
    return list_of_letters

def save_letters():
    letters = make_letters()
    file_name = "letter"
    for i in range(len(letters)):
        f = open("ReadyToSend/"+file_name + str(i)+".docx", "w")
        f.write(letters[i])

if __name__=='__main__':
    save_letters()