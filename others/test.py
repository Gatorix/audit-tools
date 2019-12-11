# # Suppose this is foo.py.

# print("before import")
# import math

# print("before functionA")
# def functionA():
#     print("Function A")

# print("before functionB")
# def functionB():
#     print("Function B {}".format(math.sqrt(100)))

# print("before __name__ guard")
# if __name__ == '__main__':
#     functionA()
#     functionB()
# print("after __name__ guard")

li = ['qqq', '111', '1qw', '222']
print(li)
n_li=[]
try:
    n_li=list(map(int,li))
except ValueError:
    
# for x in range(len(li)):
#     for y in range(len(li[x])):
#         n_li.append(li[x][y])
#         # if y == 1:
            
#         #     n_li.append(int(li[x][y]))



print(n_li)
