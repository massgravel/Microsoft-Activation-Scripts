# Windows Vista (Gunakan Edisi Business dan Enterprise)
1. Buka CMD sebagai Administrator kemudian salin dan tempel perintah ini satu persatu.
2. Jika menggunakan edisi Business

       slmgr /rilc

       slmgr /ipk YFKBB-PQJJV-G996G-VWGXY-2V3X8
   
       slmgr /skms kms8.msguides.com
   
       slmgr /ato

2. Jika menggunakan edisi Enterprise

       slmgr /rilc

       slmgr /ipk VKK3X-68KWM-X2YGT-QR4M6-4BWMV
       
       slmgr /skms kms8.msguides.com
   
       slmgr /ato
       
# Office 2010 di Windows XP/Vista
1. Buka CMD sebagai Administrator kemudian salin dan tempel perintah ini satu persatu.
   
2. Jika menggunakan OS 64-Bit:
   
       cd \Program Files (x86)\ Microsoft Office\Office14
       cscript ospp.vbs /restartosppsvc
       cscript ospp.vbs /sethst:kms8.msguides.com
       cscript ospp.vbs /act
   
3. Jika menggunakan OS 32-Bit:

       cd \Program Files\ Microsoft Office\Office14
       cscript ospp.vbs /restartosppsvc
       cscript ospp.vbs /sethst:kms8.msguides.com
       cscript ospp.vbs /act
