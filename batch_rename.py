# rename.py
# Created: 17th May 2017

'''
This will batch rename a group of files in a given direcotry
'''

__author__ = 'Cai Zhang'
__version__ = '1.0'

import os
import sys

def rename(work_dir):
    '''    
    this will batch rename a group of files in a given direcotry
    '''
    files = os.listdir(work_dir)
   # Filenames contain no path
    for filename in files:
        # Get the file extension
        file_ext = os.path.splitext(filename)[1]
        file_old = os.path.splitext(filename)[0]
        file_new = file_old.split('_')[0] + file_old.split('_')[1] + file_ext
        os.rename(os.path.join(work_dir,filename), # path as the parameters 
                  os.path.join(work_dir,file_new))


def main():
    '''
    This will be called if the script is directly invoked.
    '''
    work_dir = sys.argv[1]
    rename(work_dir)
   
    

if __name__=='__main__':
    main()
