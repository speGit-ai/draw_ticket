import os



def file_name(file_dir):
    list_name = []
    for root, dirs, files in os.walk(file_dir):
        #print(root) #当前目录路径
        #print(dirs) #当前路径下所有子目录

        for i in files:
            list_name.append(i)
            #print(i)  # 当前路径下所有非目录子文件
    return list_name

def listdir(path, list_name):  # 传入存储的list
    #filename = []

    for file in os.listdir(path):

        file_path = os.path.join(path, file)
        if os.path.isfile(file_path):   #是一个文件
            list_name.append(file_path)
            #file_name.append(file)

    #return filename

    '''
    if os.path.isdir(file_path): #是一个目录
        listdir(file_path, list_name) #递归遍历目录内的文件
    else:
        list_name.append(file_path)
    '''