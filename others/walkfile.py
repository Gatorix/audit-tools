# import os
# import shutil

# key = '试算表' # 源文件夹中的文件包含字符key则复制到to_dir_path文件夹中

# src_dir_path = 'C:\\Users\\Gatorix\\Desktop\\海能达' # 源文件夹

# to_dir_path = 'C:\\Users\\Gatorix\\Desktop\\海能达\\'+key # 存放复制文件的文件夹



# if not os.path.exists(to_dir_path):
#     print("to_dir_path not exist,so create the dir")
#     os.mkdir(to_dir_path,1)

# if os.path.exists(src_dir_path):
#     print("src_dir_path exitst")

# for file in os.listdir(src_dir_path):

#     if os.path.isfile(src_dir_path+'/'+file):

#         if key in file:

#             print('找到包含"'+key+'"字符的文件,绝对路径为----->'+src_dir_path+'/'+file)
#             print('复制到----->'+to_dir_path+file)
#             shutil.copy(src_dir_path+'/'+file,to_dir_path+'/'+file)
import os
# key = '试算表' # 源文件夹中的文件包含字符key则复制到to_dir_path文件夹中

# src_dir_path = 'C:\\Users\\Gatorix\\Desktop\\海能达' # 源文件夹

# to_dir_path = 'C:\\Users\\Gatorix\\Desktop\\海能达\\'+key # 存放复制文件的文件夹

path = 'C:\\Users\\Gatorix\\Desktop\\海能达'

for dirpath,dirnames,filenames in os.walk(path):

    for filename in filenames:

        print(os.path.join(dirpath,filename))
