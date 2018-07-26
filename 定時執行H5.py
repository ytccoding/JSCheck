# -*- coding: utf-8 -*-

import subprocess
from subprocess import PIPE
from time import sleep

startNumber = input("開始號碼(從1開始):").strip()
endNumber = input("結束號碼:").strip()

    
for i in range(int(startNumber) ,int(endNumber)+1):
   p = subprocess.Popen('JS檢查_chrome_H5_1.3.py' ,stdout=PIPE ,stderr=PIPE ,stdin=PIPE ,shell = True)
   try:
      p.communicate(input=str(i).encode("utf-8")  ,timeout = 2)
   except:
      pass
   sleep(2)


