# Windows Vista (Edisi Business, Business N, Enterprise)
1. Pastikan perangkat memiliki akses internet 
2. Buka CMD sebagai Administrator kemudian salin dan tempel perintah ini satu persatu.
3. Masukkan [Generic Product Key](https://learn.microsoft.com/id-id/windows-server/get-started/kms-client-activation-keys?tabs=server2025%2Cwindows1110ltsc%2Cversion1803%2Cwindows81) Sesuai Edisi pada perintah **slmgr /ipk**

| Edisi | Product Key |
| --- | --- |
| Windows Vista Business | YFKBB-PQJJV-G996G-VWGXY-2V3X8 |
| Windows Vista Business N | HMBQG-8H2RH-C77VX-27R82-VMQBT |
| Windows Vista Enterprise | VKK3X-68KWM-X2YGT-QR4M6-4BWMV |
| Windows Vista Enterprise N | VTC42-BM838-43QHV-84HX6-XJXKV |

       slmgr /rilc

       slmgr /ipk <Product Key>
   
       slmgr /skms kms8.msguides.com
   
       slmgr /ato

# Windows Vista (Edisi lainnya, contoh: Ultimate, HomePremium, HomePremium N, HomeBasic, HomeBasic N, Starter)
Gunakan [Windows Loader by Daz](https://app.box.com/s/bnchc6hten44adunlcpz9ya9j0uucfs2)
       
# Office 2010 di Windows XP/Vista
1. Pastikan perangkat memiliki akses internet 
2. Buka CMD sebagai Administrator kemudian salin dan tempel perintah ini satu persatu.
   
       cd \Program Files\Microsoft Office\Office14
   
       cscript ospp.vbs /restartosppsvc
   
       cscript ospp.vbs /sethst:kms8.msguides.com
   
       cscript ospp.vbs /act
   
**Catatan: Jika menggunakan OS 64-Bit maka perintah pertama adalah:**

       cd \Program Files (x86)\Microsoft Office\Office14
