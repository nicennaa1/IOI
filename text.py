# version:python3.7.0
# Character：UTF-8
# AUTHOR:LUO DI
import pypyodbc
# import mailmerge
from mailmerge import MailMerge

template = 'C://Users/21452/Desktop/sc/1.docx'
url = r'Driver={Microsoft Access Driver (*.mdb,*.accdb)};DBQ=C:\Users\21452\Desktop\sc\1\3208302022.mdb'
conn = pypyodbc.win_connect_mdb(url)
cur = conn.cursor()
#cur.execute("""SELECT M.NHDM,M.HZXM,M.HZZJHM,M.TXDZ,M.HNCYSL,M.SFWBH,M.SFPKH,M.SJLY,(SELECT N.XM + '|' + N.ZJLX + '|' + N.YHZGX + '|' + N.ZJHM + '|' + N.SFBJTJJZZCY + '|' + N.XB + '|' + N.HKLX + '|' + N.HYZK + '|' + N.LXDH + '|'  FROM NHHNCY N WHERE M.NHDM = N.NHDM FOR XML PATH('')) AS MEMBERS FROM NH M GROUP BY M.NHDM,M.HZXM,M.HZZJHM,M.TXDZ,M.HNCYSL,M.SFWBH,M.SFPKH,M.SJLY ORDER BY M.NHDM""")
#cur.execute("""SELECT M.NHDM,M.HZXM,M.HZZJHM,M.TXDZ,M.HNCYSL,M.SFWBH,M.SFPKH,M.SJLY,N.XM,N.ZJLX,N.YHZGX,N.ZJHM,N.SFBJTJJZZCY,N.XB,N.HKLX,N.HYZK,N.LXDH FROM NH AS M INNER JOIN NHHNCY AS N ON M.NHDM = N.NHDM GROUP BY M.NHDM ORDER BY M.NHDM""")
#cur.execute("""SELECT M.NHDM,MAX(M.HZXM),MAX(M.HZZJHM),MAX(M.TXDZ),MAX(M.HNCYSL),MAX(M.SFWBH),MAX(M.SFPKH),MAX(M.SJLY),MAX(N.XM),MAX(N.ZJLX),MAX(N.YHZGX),MAX(N.ZJHM),MAX(N.SFBJTJJZZCY),MAX(N.XB),MAX(N.HKLX),MAX(N.HYZK),MAX(N.LXDH) FROM NH AS M INNER JOIN NHHNCY AS N ON M.NHDM = N.NHDM GROUP BY M.NHDM ORDER BY M.NHDM""")
cur.execute("""SELECT M.NHDM, FIRST(M.HZXM) AS HZXM_FIRST, LAST(M.HZXM) AS HZXM_LAST, FIRST(M.HZZJHM) AS HZZJHM_FIRST, LAST(M.HZZJHM) AS HZZJHM_LAST, 
       FIRST(M.TXDZ) AS TXDZ_FIRST, LAST(M.TXDZ) AS TXDZ_LAST, FIRST(M.HNCYSL) AS HNCYSL_FIRST, LAST(M.HNCYSL) AS HNCYSL_LAST, 
       FIRST(M.SFWBH) AS SFWBH_FIRST, LAST(M.SFWBH) AS SFWBH_LAST, FIRST(M.SFPKH) AS SFPKH_FIRST, LAST(M.SFPKH) AS SFPKH_LAST, 
       FIRST(M.SJLY) AS SJLY_FIRST, LAST(M.SJLY) AS SJLY_LAST, 
       FIRST(N.XM) AS XM_FIRST, LAST(N.XM) AS XM_LAST, FIRST(N.ZJLX) AS ZJLX_FIRST, LAST(N.ZJLX) AS ZJLX_LAST, 
       FIRST(N.YHZGX) AS YHZGX_FIRST, LAST(N.YHZGX) AS YHZGX_LAST, FIRST(N.ZJHM) AS ZJHM_FIRST, LAST(N.ZJHM) AS ZJHM_LAST, 
       FIRST(N.SFBJTJJZZCY) AS SFBJTJJZZCY_FIRST, LAST(N.SFBJTJJZZCY) AS SFBJTJJZZCY_LAST, FIRST(N.XB) AS XB_FIRST, LAST(N.XB) AS XB_LAST, 
       FIRST(N.HKLX) AS HKLX_FIRST, LAST(N.HKLX) AS HKLX_LAST, FIRST(N.HYZK) AS HYZK_FIRST, LAST(N.HYZK) AS HYZK_LAST, 
       FIRST(N.LXDH) AS LXDH_FIRST, LAST(N.LXDH) AS LXDH_LAST
FROM NH M
LEFT JOIN NHHNCY N ON M.NHDM = N.NHDM
GROUP BY M.NHDM""")

result = cur.fetchall()
print(len(result))
row = 0
for i in result:
     document_new = MailMerge(template)
     print('{}文档中，邮件合并的字段包含:{}'.format(template, document_new.get_merge_fields()))
     str_list = list(map(str,i))
     x=len(str_list)
     #while len(str_list)<200:
         #str_list.append([None])
     if len(str_list) >= 100:
        str_list[100] = None
     else:
         some_value=None
         str_list += [None] * (100 - len(str_list)) + [None]
         print(str_list[19])
     data = document_new.merge(NHDM = str_list[0],
                         HZXM = str_list[1],
                         HZZJHM = str_list[2],
                         TXDZ = str_list[3],
                         HNCYSL = str_list[4],
                         SFWBH = str_list[5],
                         SFPKH = str_list[6],
                         SJLY = str_list[7],
                         XM1 = str_list[8],
                         ZJLX1 = str_list[9],
                         YHZGX1 = str_list[10],
                         ZJHM1 = str_list[11],
                         SFBJTJJZZCY1=str_list[12],
                         XB1 = str_list[13],
                         HKLX1 = str_list[14],
                         HYZK1 = str_list[15],
                         HYZK=str_list[15],
                         LXDH1 = str_list[16],
                         LXDH=str_list[16],
                         XM2 = str_list[17],
                         ZJLX2 = str_list[18],
                         YHZGX2 = str_list[19],
                         ZJHM2 = str_list[20],
                         SFBJTJJZZCY2 = str_list[21],
                         XB2 = str_list[22],
                         HKLX2 = str_list[23],
                         HYZK2=str_list[24],
                         LXDH2=str_list[25],
                         XM3=str_list[17],
                         ZJLX3=str_list[18],
                         YHZGX3=str_list[19],
                         ZJHM3=str_list[20],
                               )
     row += 1
     document_new.write(r'C:\Users\21452\Desktop\sc\JG\地籍调查报告-{}.docx'.format(str_list[1]))

     #我要的是 代码、 户主、手机号 、互助成员、手机号、
     #       1000、 张三、 132、    张三、   张四、 123、 123
     #       1001、 李四、 123、    李武、    123 、  李六、  123 、李七、 123

     #       1000、  张三、 123 、张四、123
     #       1001、  李四、 123 、李五、123


