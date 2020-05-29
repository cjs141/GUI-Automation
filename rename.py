import os

os.chdir('/Users/swordfish/Public/TXST/seSP20/testRename')

for f in os.listdir('/Users/swordfish/Public/TXST/seSP20/testRename'):
    f_name, f_ext = (os.path.splitext(f))
    f_title, f_place, f_num = f_name.split('-')

    f_title = f_title.strip()
    f_place = f_place.strip()
    f_num = f_num.strip()[1:].zfill(2) 

    new_name = ('{}-{}{}'.format(f_num, f_title, f_ext))

    os.rename(f, new_name)






