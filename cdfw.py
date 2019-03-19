#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pandas as pd
from datetime import datetime
import numpy as np
import matplotlib.pyplot as plt
plt.rcParams['font.sans-serif']=['SimHei'] #用来显示中文
plt.rcParams['axes.unicode_minus']=False #用来正常显示负号
import re,random
from tkinter import *
import tkinter.filedialog


# In[2]:


#辅助函数
def get_check_type(x):#check提前检查纯数字、日期等不规范的记法,并给出解决办法
    if isinstance(x,pd.Timestamp) or isinstance(x,pd.datetime):
        return "错成excel时间格式了，请在音节前加'/'，如'/jan22'"
    elif isinstance(x,int):
        return "错成纯数字了"
    elif '\\' in x:
        return "请不要用'\\',改用'/'"
    elif '/' in x.strip('/'):
        return "请不要在音节中用'/'，如一词声母有变体，请拆成两个条目"
    else:
        return ""
def init(x):
    return x
def init_examples(x,ned,example_num):
    return ' / '.join(sorted(list(ned.loc[ned['init']==x,'item'].unique()),key=len)[:example_num])

def fnl_phon_sort(arr,fnl_num_dict):
    for i in range(len(arr)):
        for j in range(i+1,len(arr)):
            if fnl_num_dict[arr[i]]>fnl_num_dict[arr[j]]:
                arr[i],arr[j]=arr[j],arr[i]
def fnl_num(fnl,medials,fnl_sect_dict):
    num=[0,0,0,0]
    num[0]=medials.index(fnl_sect_dict[fnl][0])
    num[0]=0 if num[0]==-1 else num[0]
    num[1]=vowel_num(fnl_sect_dict[fnl][1]+fnl_sect_dict[fnl][2])*10
    num[2]=vowel_num(fnl_sect_dict[fnl][3])*100
    num[3]=["","m","n","ng","p","t","k","?"].index(fnl_sect_dict[fnl][4])*10000000
    return sum(num)
def vowel_num(vowel):
    if vowel=="":
        return 0
    if vowel in ["z","v","l","m","n","ng"]:
        pos=["z","v","l","m","n","ng"].index(vowel)
        return pos*1000000
    ipa_fnl_str ='''i~,y~,i#~,u#~,u=~,u~;I~,Y~,,,U~,;e~,e@~,e#~,o#~,e>~,o~;E~,,e=~,,,o=~;e+~,e+@~,e+#~,o+#~,o+$~,o+~;a^~,,A^~,,,;a~,a@~,A~,,a>~,a>@~;i<~,i>~,y<~,y>~,,;
    i,y,i#,u#,u=,u;I,Y,,,U,;e,e@,e#,o#,e>,o;E,,e=,,,o=;e+,e+@,e+#,o+#,o+$,o+;a^,,A^,,,;a,a@,A,,a>,a>@;i<,i>,y<,y>,,;
    i:,y:,i#:,u#:,u=:,u:;I:,Y:,,,U:,;e:,e@:,e#:,o#:,e>:,o:;E:,,e=:,,,o=:;e+:,e+@:,e+#:,o+#:,o+$:,o+:;a^:,,A^:,,,;a:,a@:,A:,,a>:,a>@:;i<:,i>:,y<:,y>:,,'''
    ipa_fnl=ipa_fnl_str.split(";")
    for i,v in enumerate(ipa_fnl):
        if vowel in v.split(","):
            pos=v.split(",").index(vowel)
            num=(len(ipa_fnl)-i)*10+pos
            if vowel[-1]=="~":
                num*=100
            return num
    return random.randint(10000,20000)    

def get_multi_tones(x):
    xs=x.split(' ')
    for k in range(len(xs)):
        for i in range(len(xs[k])-1,0,-1):
            if xs[k][i] not in '0123456789':
                if xs[k][i] in '?ptk':
                    xs[k]=xs[k][i+1:]+'入'
                else:xs[k]=xs[k][i+1:]  
                break
    return ' '.join(xs)


# In[3]:


def part_execute():
    global DATA_READY
    if not DATA_READY:
        note_lb.config(text="源数据没有选择对，请重新选择！")
    else:
        part_analyse()

def part_analyse():
    #try:
        global SOURCE
        global TARGETSHEET
        data=pd.read_excel(SOURCE,sheet_name=TARGETSHEET.get())
        #清洗
        #列名重命名
        init_col=data.columns
        now_col=[""]*len(init_col)
        for i,col in enumerate(init_col):
            if re.match(u".*[词项|汉字|词目|词|字].*",col):
                now_col[i]="item"
            elif re.match(u".*记音.*",col):
                now_col[i]="rec"
            elif re.match(u".*文白.*",col):
                now_col[i]="wenbai"
            elif re.match(u".*义.*",col):
                now_col[i]="meaning"
            elif re.match(u".*备注.*",col):
                now_col[i]="note"
            else:
                now_col[i]=col
        col_dict={init_col[i]:now_col[i] for i in range(len(init_col))}
        data.rename(columns=col_dict,inplace=True)
        #去除空白项
        data=data.fillna("")

        data["check"]=data["rec"].apply(get_check_type)
        data["rec"]=data["rec"].apply(lambda x:str(x))#把记录全部变为字符串
        data=data[data["rec"]!=""]#去掉空行
        #去除重复项,直接删了
        data=data.drop_duplicates(["item","rec"])
        data.sort_values(by=["item","rec"],ascending=False)
        data=data.reset_index(drop=True)#重新排索引，把中间之前删掉的索引补回来
        #增添新列
        extra_col=now_col+["check","std_syl","init","center","tail","tone",'orig']
        data=data.reindex(columns=extra_col,fill_value="")
        data.loc[:,"sylcnt"]=0
        data.loc[:,"orig"]='原数据'
        #规范记音,切分多音节，标注，切分声韵调
        #判断所用的字符
        fnl_sect_dict={}#后面做韵母的语音学排序会用到
        std_letters="qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM0123456789~!?@#$^&<>*=+:"
        for idx,syl in enumerate(data["rec"]):
            new_syl=""
            for i,letter in enumerate(syl):
                if letter in std_letters:
                    new_syl+=letter
            #正则切分
            pattern=re.compile("(\??)([qwrtypsdfghjklzxcvbnmQWRTYPSDFGHJKLZXCVBNM>#\!\$\^\*\+]{0,4})([aeiouAEIOUw]?[<>=#~@:\$\^\*\+]*)([aeiouAEIOUmnlzv]|ng){1}([<>=#~@:\$\^\*\+]*)([aeiouAEIOU]?[<>=#~@:\$\^\*\+]*)([mnptk\?]|ng)?(\d+)")
            results=pattern.findall(new_syl)
            for i in range(len(results)):#细节调整
                results[i]=list(results[i])
                result=results[i]
                if result[0]+result[1] in ["",'w']:#零声母的情况，默认补一个？，避免'?'与''重复计数
                    result[0]="?"
                if len(result[1])>0 and result[1][-1]=='w':
                    result[1]=result[1][:-1]
                    result[2]='w'+result[2]#w算作介音
                if result[2]!="" and result[3] in ["n","ng","m"]:
                    result[6]=result[3]
                    result[3]=result[2]
                    result[2]=""
                elif result[3] in ["n","ng","m"] and result[6] in ["n","ng","m"]:
                    result[3]=""
                if result[6] in ["p","t","k","?"]:
                    result[7]+="入"
                if result[2]!="" and result[3] in "iuyIUY" and result[5]=="":
                    result[5]=result[3]+result[4]
                    result[4]=""
                    result[3]=result[2]
                    result[2]=""

            #切分声韵调
            if len(results)==1:
                sylsects=results[0]
                data.loc[idx,"std_syl"]="".join(filter(None,sylsects)).strip("入")
                data.loc[idx,"init"]=sylsects[0]+sylsects[1]
                data.loc[idx,"center"]=sylsects[2]+sylsects[3]+sylsects[4]+sylsects[5]
                data.loc[idx,"tail"]=sylsects[6]
                data.loc[idx,"tone"]=sylsects[7]
                data.loc[idx,"sylcnt"]=1
                fnl=sylsects[2]+sylsects[3]+sylsects[4]+sylsects[5]+sylsects[6]
                if fnl not in fnl_sect_dict:
                    fnl_sect_dict[fnl]=[sylsects[2],sylsects[3],sylsects[4],sylsects[5],sylsects[6]]

            elif len(results)>1:
                data.loc[idx,"sylcnt"]=len(results)
                poly_sects=[]
                for j,result in enumerate(results):
                    row={data.columns[i]:"" for i in range(len(data.columns))}
                    row['rec']=data.loc[idx,'rec']
                    row["item"]=data.loc[i,"item"]+str(j+1)
                    tmp_syl="".join(filter(None,result))
                    tmp_syl=tmp_syl.strip("入")
                    poly_sects.append(tmp_syl)
                    row["std_syl"]=tmp_syl
                    row["init"]=result[0]+result[1]
                    row["center"]=result[2]+result[3]+result[4]+result[5]
                    row["tail"]=result[6]
                    row["tone"]=result[7]
                    row["sylcnt"]=1
                    row['orig']='切分出的单音节'
                    data=data.append(row,ignore_index=True)
                    fnl=result[2]+result[3]+result[4]+result[5]+result[6]
                    if fnl not in fnl_sect_dict:
                        fnl_sect_dict[fnl]=[result[2],result[3],result[4],result[5],result[6]]
                data.loc[idx,"std_syl"]=" ".join(poly_sects)
        #统计多音字/词
        item_cnt=data.groupby(["item"],as_index=False)["item"].agg({"ppcnt":"count"})
        #item_cnt=data["item"].value_counts()
        data=pd.merge(data,item_cnt,left_on=["item"],right_on=["item"],how="left")
        #程序标注需要修改的项
        data.loc[data["rec"]!=data["std_syl"],"check"]+="  记音不规范"
        data.loc[data['orig']=='切分出的单音节',"check"]=''#对于多音节词拆分出来的单音节不需要检查，直接检查多音节就好了

        #要输出的结果表
        nowstr=datetime.now().strftime("%m%d")
        save_path=SOURCE.split('.xls')[0]+'切分结果'+nowstr+".xlsx"
        writer=pd.ExcelWriter(save_path)
        data.to_excel(writer,sheet_name='切分表',index=False,header=True)

        #联系作者
        add='https://github.com/Lykit01/Lykit-for-chinese-dialects-field-work'
        about=pd.Series(['@Hue Zhang','3275803255@qq.com',add,'如有问题请联系我！'],index=['author','email','github',''])
        about.to_excel(writer,sheet_name='关于')
        
        writer.save()
        note_lb.config(text="分析完毕，结果保存在:"+save_path)
    #except Exception:
        #note_lb.config(text="表内数据有误！请检查源数据！")


# In[4]:


def execute():
    global DATA_READY
    if not DATA_READY:
        note_lb.config(text="源数据没有选择对，请重新选择！")
    else:
        analyse()

#分析程序
def analyse():
    #try:
        global SOURCE
        global TARGETSHEET
        global EXAMPLE_NUM
        global INIT_EXCEP
        global FNL_EXCEP
        global TONE_EXCEP
        excep_val=[int(INIT_EXCEP.get()),int(FNL_EXCEP.get()),int(TONE_EXCEP.get())]
        example_num=int(EXAMPLE_NUM.get())
        data=pd.read_excel(SOURCE,sheet_name=TARGETSHEET.get())
        #清洗
        #列名重命名
        init_col=data.columns
        now_col=[""]*len(init_col)
        for i,col in enumerate(init_col):
            if re.match(u".*[词项|汉字|词目|词|字].*",col):
                now_col[i]="item"
            elif re.match(u".*记音.*",col):
                now_col[i]="rec"
            elif re.match(u".*文白.*",col):
                now_col[i]="wenbai"
            elif re.match(u".*义.*",col):
                now_col[i]="meaning"
            elif re.match(u".*备注.*",col):
                now_col[i]="note"
            else:
                now_col[i]=col
        col_dict={init_col[i]:now_col[i] for i in range(len(init_col))}
        data.rename(columns=col_dict,inplace=True)
        #去除空白项
        data=data.fillna("")

        data["check"]=data["rec"].apply(get_check_type)
        data["rec"]=data["rec"].apply(lambda x:str(x))#把记录全部变为字符串
        data=data[data["rec"]!=""]#去掉空行
        #去除重复项,直接删了
        data=data.drop_duplicates(["item","rec"])
        data.sort_values(by=["item","rec"],ascending=False)
        data=data.reset_index(drop=True)#重新排索引，把中间之前删掉的索引补回来
        #增添新列
        extra_col=now_col+["check","std_syl","init","center","tail","tone",'orig']
        data=data.reindex(columns=extra_col,fill_value="")
        data.loc[:,"sylcnt"]=0
        data.loc[:,"orig"]='原数据'
        #规范记音,切分多音节，标注，切分声韵调
        #判断所用的字符
        fnl_sect_dict={}#后面做韵母的语音学排序会用到
        std_letters="qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM0123456789~!?@#$^&<>*=+:"
        for idx,syl in enumerate(data["rec"]):
            new_syl=""
            for i,letter in enumerate(syl):
                if letter in std_letters:
                    new_syl+=letter
            #正则切分
            pattern=re.compile("(\??)([qwrtypsdfghjklzxcvbnmQWRTYPSDFGHJKLZXCVBNM>#\!\$\^\*\+]{0,4})([aeiouAEIOUw]?[<>=#~@:\$\^\*\+]*)([aeiouAEIOUmnlzv]|ng){1}([<>=#~@:\$\^\*\+]*)([aeiouAEIOU]?[<>=#~@:\$\^\*\+]*)([mnptk\?]|ng)?(\d+)")
            results=pattern.findall(new_syl)
            for i in range(len(results)):#细节调整
                results[i]=list(results[i])
                result=results[i]
                if result[0]+result[1] in ["",'w']:#零声母的情况，默认补一个？，避免'?'与''重复计数
                    result[0]="?"
                if len(result[1])>0 and result[1][-1]=='w':
                    result[1]=result[1][:-1]
                    result[2]='w'+result[2]#w算作介音
                if result[2]!="" and result[3] in ["n","ng","m"]:
                    result[6]=result[3]
                    result[3]=result[2]
                    result[2]=""
                elif result[3] in ["n","ng","m"] and result[6] in ["n","ng","m"]:
                    result[3]=""
                if result[6] in ["p","t","k","?"]:
                    result[7]+="入"
                if result[2]!="" and result[3] in "iuyIUY" and result[5]=="":
                    result[5]=result[3]+result[4]
                    result[4]=""
                    result[3]=result[2]
                    result[2]=""

            #切分声韵调
            if len(results)==1:
                sylsects=results[0]
                data.loc[idx,"std_syl"]="".join(filter(None,sylsects)).strip("入")
                data.loc[idx,"init"]=sylsects[0]+sylsects[1]
                data.loc[idx,"center"]=sylsects[2]+sylsects[3]+sylsects[4]+sylsects[5]
                data.loc[idx,"tail"]=sylsects[6]
                data.loc[idx,"tone"]=sylsects[7]
                data.loc[idx,"sylcnt"]=1
                fnl=sylsects[2]+sylsects[3]+sylsects[4]+sylsects[5]+sylsects[6]
                if fnl not in fnl_sect_dict:
                    fnl_sect_dict[fnl]=[sylsects[2],sylsects[3],sylsects[4],sylsects[5],sylsects[6]]

            elif len(results)>1:
                data.loc[idx,"sylcnt"]=len(results)
                poly_sects=[]
                for j,result in enumerate(results):
                    row={data.columns[i]:"" for i in range(len(data.columns))}
                    row['rec']=data.loc[idx,'rec']
                    row["item"]=data.loc[i,"item"]+str(j+1)
                    tmp_syl="".join(filter(None,result))
                    tmp_syl=tmp_syl.strip("入")
                    poly_sects.append(tmp_syl)
                    row["std_syl"]=tmp_syl
                    row["init"]=result[0]+result[1]
                    row["center"]=result[2]+result[3]+result[4]+result[5]
                    row["tail"]=result[6]
                    row["tone"]=result[7]
                    row["sylcnt"]=1
                    row['orig']='切分出的单音节'
                    data=data.append(row,ignore_index=True)
                    fnl=result[2]+result[3]+result[4]+result[5]+result[6]
                    if fnl not in fnl_sect_dict:
                        fnl_sect_dict[fnl]=[result[2],result[3],result[4],result[5],result[6]]
                data.loc[idx,"std_syl"]=" ".join(poly_sects)
        #统计多音字/词
        item_cnt=data.groupby(["item"],as_index=False)["item"].agg({"ppcnt":"count"})
        #item_cnt=data["item"].value_counts()
        data=pd.merge(data,item_cnt,left_on=["item"],right_on=["item"],how="left")
        #程序标注需要修改的项
        data.loc[data["rec"]!=data["std_syl"],"check"]+="  记音不规范"
        data.loc[data['orig']=='切分出的单音节',"check"]=''#对于多音节词拆分出来的单音节不需要检查，直接检查多音节就好了
        #去除没有切分的多音节词
        #ssd:single syl data
        ssd=data[data["center"]!=""].copy()#这里用副本，避免下面对ssd进行筛选时触发复制警告
        #包含例外的音节
        whole_syl_list=ssd["std_syl"].value_counts()
        #声韵组合分析
        #init_and_fnl_list=ssd.groupby(["init","center","tail"])["item"].count().sort_values(ascending=False)
        init_and_fnl_list=(ssd.loc[:,"init"]+ssd.loc[:,"center"]+ssd.loc[:,"tail"]).value_counts()
        #包括例外的声母
        init_list=ssd["init"].value_counts()
        #包含例外的韵母
        ssd.loc[:,"fnl"]=ssd.loc[:,"center"]+ssd.loc[:,"tail"]#两列直接相加要改为这种
        fnl_list=ssd["fnl"].value_counts()
        #包含例外的声调
        tone_list=ssd["tone"].value_counts()
        tail_list=ssd["tail"].value_counts()

        #不包括例外的声母
        no_excep_init=init_list[init_list>excep_val[0]]
        #不包含例外的韵母
        no_excep_fnl=fnl_list[fnl_list>excep_val[1]]
        #不包含例外的声调
        no_excep_tone=tone_list[tone_list>excep_val[2]]

        #no_excep_data
        ned=ssd[ssd["init"].isin(no_excep_init.index)&ssd["fnl"].isin(no_excep_fnl.index)&ssd["tone"].isin(no_excep_tone.index)].copy()
        #不包含例外的音节
        no_excep_whole_syl=ned["std_syl"].value_counts()
        #不包含例外的声韵组合个数
        no_excep_init_and_fnl=(ned.loc[:,"init"]+ned.loc[:,"fnl"]).value_counts()
        #不包含例外的韵尾个数
        no_excep_tail=ned["tail"].value_counts()

        #对含例外的记录标上记号"例外"，回到手工检查这一步
        data.loc[data["sylcnt"]==1&-data["std_syl"].isin(no_excep_whole_syl.index),"check"]+="声韵调例外"
        survey=pd.Series({"记录的单音节数":data.loc[data["sylcnt"]==1,"std_syl"].count(),
                      "记录的多音节词数":data.loc[data["sylcnt"]>1,"std_syl"].count(),
                     "拆分多音节得到的单音节":data.loc[data["rec"]=="","std_syl"].count(),
                      "多音字/词的个数":len(data.loc[(data["ppcnt"]>1) & (data["rec"]!=""),"item"].unique()),
                    "音节数(去重,含例外)":len(whole_syl_list),
                      "声韵组合数(去重,含例外)":len(init_and_fnl_list),
                      "声母数(去重,含例外)":len(init_list),
                      "韵母数(去重,含例外)":len(fnl_list),
                      "声调数(去重,含例外)":len(tone_list),
                      "辅音韵尾数(去重,含例外)":len(tail_list),
                   "音节数(去重,不含例外)":len(no_excep_whole_syl),
                      "声韵组合数(去重,不含例外)":len(no_excep_init_and_fnl),
                      "声母数(去重,不含例外)":len(no_excep_init),
                      "韵母数(去重,不含例外)":len(no_excep_fnl),
                      "声调数(去重,不含例外)":len(no_excep_tone),
                      "辅音韵尾数(去重,不含例外)":len(no_excep_tail)
                    })
        if len(ned)>0:
            #语音学排序
            #声母Series
            init_tmp=list(no_excep_init.index)#排除例外后的声母列表
            init_phon=[]#最后的声母列表
            init_df=pd.DataFrame(init_tmp,index=init_tmp)
            ipa_init_part_list=["双唇","唇齿","舌齿","舌尖前","舌尖中","舌尖后","舌叶","舌面前","舌面中","舌面后","小舌","喉壁","喉门"]
            ipa_init_manner_list=["不送清塞","送气清塞","不送浊塞","送气浊塞","不送清塞擦","送气清塞擦","不送浊塞擦","送气浊塞擦","鼻","颤音","闪音","边","清边擦","浊边擦","清擦","浊擦","通展","通圆"]
            ipa_init_str = "p,ph,b,bh,,,,,m,,,,,,p*,b*,,;,,,,pf,pfh,bv,bvh,mg,,,,,,f,v,v$,;,,,,t>,t>h,d>,d>h,,,,,,,s>,z>,,;,,,,ts,tsh,dz,dzh,,,,,,,s,z,,;t,th,d,dh,,,,,n,r,r*,l,ls,l#,,,r$,;tr,trh,dr,drh,tsr,tsrh,dzr,dzrh,nr,,r^,lr,,,sr,zr,rr,;,,,,tss,tssh,dzz,dzzh,,,,,,,ss,zz,,;tj,tjh,dj,djh,tcj,tcjh,dzj,dzjh,nj,,,,,,cj,zj,,;c,ch,c!,c!h,,,,,nc,,,lc,,,c#,jj,j,y$;k,kh,g,gh,,,,,ng,,,,,,x,x!,w!,w$;q,qh,G,Gh,,,,,N,R,,,,,X,X!,,;,,,,,,,,,,,,,,h*,h*!,,;?,?h,,,,,,,,,,,,,h,h!,,"
            ipa_init=[s.split(",") for s in ipa_init_str.split(";")]
            for i in range(len(ipa_init)):
                for j in range(len(ipa_init[0])):
                    letter=ipa_init[i][j]
                    if letter!='' and letter in init_tmp:
                        init_df.loc[letter,'发音部位']=ipa_init_part_list[i]
                        init_df.loc[letter,'发音方法']=ipa_init_manner_list[j]
                        init_phon.append(letter)
            #不在标准国际音标表中的声母        
            init_phon+=list(init_df.loc[init_df['发音方法'].isna(),0])
            init_df.loc[init_df['发音部位'].isna(),'发音部位']=init_df.loc[:,0]
            init_df.loc[init_df['发音方法'].isna(),'发音方法']='其他'
            init_table=init_df.groupby(['发音部位','发音方法'])[0].sum().unstack(1).reindex(index=ipa_init_part_list+init_tmp,columns=ipa_init_manner_list+['其他'])
            init_table=init_table.dropna(how='all').dropna(how='all',axis=1).fillna('')
            #含例字的声母表
            init_table_with_examples=init_table.transform([init,init_examples],ned=ned,example_num=example_num)
            #韵母

            fnl_phon=list(no_excep_fnl.index)#fnl转换格式
            #处理介音问题
            medials=set()#介音
            for fnl in fnl_phon:
                medials.add(fnl_sect_dict[fnl][0])
            medials=list(medials)
            medials.sort()
            #韵母排序
            fnl_num_dict={}
            for fnl in fnl_phon:
                fnl_num_dict[fnl]=fnl_num(fnl,medials,fnl_sect_dict)
            fnl_phon_sort(fnl_phon,fnl_num_dict)
            for i in range(len(fnl_phon)-1,0,-1):
                fnl=fnl_phon[i]
                if fnl_sect_dict[fnl][4] not in ['p','t','k','?']:
                    break
            fnl_phon_shu=fnl_phon[:i+1]
            fnl_phon_cu=fnl_phon[i+1:]
            #韵母表Table 20190317
            fnl_table_list=[]
            fnl_articulation_dict={}
            last_dval=0
            fnl_base=[]#韵基
            for fnl in fnl_phon:
                dval=fnl_num_dict[fnl]//10
                if dval!=last_dval:
                    last_dval=dval
                    fnl_base.append(fnl)
                    fnl_table_list.append(["" for _ in medials])
                mval=fnl_num_dict[fnl]%10
                fnl_table_list[-1][mval]=fnl
                fnl_articulation_dict[fnl]=(fnl_base[-1],medials[mval])
            fnl_table=pd.DataFrame(fnl_table_list,index=fnl_base,columns=medials)
            #含例字的韵母表
            new_index=[]
            for x in fnl_base:
                new_index+=[x]+[x+'例字']
            fnl_table_with_examples=fnl_table.reindex(index=new_index).fillna('')
            for fnl in fnl_phon:
                options=list(ned.loc[ned['fnl']==fnl,'item'].unique())
                fnl_table_with_examples.loc[fnl_articulation_dict[fnl][0]+'例字',
                    fnl_articulation_dict[fnl][1]]=' / '.join(sorted(options,key=len)[:example_num])
            #声调
            tone_phon=list(no_excep_tone.index)
            tone_phon.sort(key=lambda x:x[-1])
            for i in range(len(tone_phon)-1,0,-1):
                if tone_phon[i][-1]!="入":break
            tone_phon_shu=tone_phon[:i+1]
            tone_phon_cu=tone_phon[i+1:]
            tone_phon=pd.DataFrame(tone_phon,index=tone_phon,columns=['调值'])
            tone_phon['例字']=tone_phon['调值'].apply(lambda x:' / '.join(sorted(list(ned.loc[ned['tone']==x,'item'].unique()),key=len)[:example_num]))
            #同音字表 新版20190318
            hpp_tuples=list(zip([x for x in fnl_phon_shu for i in range(len(tone_phon_shu))]+[x for x in fnl_phon_cu for i in range(len(tone_phon_cu))],
                            tone_phon_shu*len(fnl_phon_shu)+tone_phon_cu*len(fnl_phon_cu)))
            hpp_index=pd.MultiIndex.from_tuples(hpp_tuples,names=["韵母","声调"])
            #homophone_table=ned.groupby(['fnl','tone','init'])['item'].apply(lambda x:' // '.join(x))
            #homophone_count_table=ned.groupby(['fnl','tone','init'])['item'].count()
            homophone_all=ned.groupby(['fnl','tone','init'])['item'].apply(lambda x:str(x.count())+' // '+' // '.join(x))
            homophone_all=homophone_all.unstack(2).reindex(index=hpp_index,columns=init_phon)
            homophone_all=homophone_all.fillna('')
            #调型
            tone_match_init_fnl=ned.groupby(['fnl','init'])['tone'].apply(lambda x:' / '.join(x))
            tone_match_init_fnl=tone_match_init_fnl.unstack(1).reindex(index=fnl_phon,columns=init_phon)
            tone_match_init_fnl=tone_match_init_fnl.fillna('')
        #多音节声调组合
        #二音节声调组合
        two_syl_data=data.loc[data['sylcnt']==2,:].copy()
        if len(two_syl_data)>0:
            two_syl_data.loc[:,'itemrec']=two_syl_data.loc[:,'item']+': '+two_syl_data.loc[:,'std_syl']
            two_syl_data.loc[:,'two_tones']=two_syl_data.loc[:,'std_syl'].apply(get_multi_tones)
            two_syl_data.loc[:,'firsttone']=two_syl_data.loc[:,'two_tones'].apply(lambda x:x.split(' ')[0])
            two_syl_data.loc[:,'secondtone']=two_syl_data.loc[:,'two_tones'].apply(lambda x:x.split(' ')[1])
            two_syl_table=two_syl_data.groupby(['firsttone','secondtone'])['itemrec'].apply(lambda x:' // '.join(x))
            two_syl_table=two_syl_table.unstack(1)
        #三音节声调组合
        tri_syl_data=data.loc[data['sylcnt']==3,:].copy()
        if len(tri_syl_data)>0:
            tri_syl_data.loc[:,'itemrec']=tri_syl_data.loc[:,'item']+': '+tri_syl_data.loc[:,'std_syl']
            tri_syl_data.loc[:,'tri_tones']=tri_syl_data.loc[:,'std_syl'].apply(get_multi_tones)
            tri_syl_data.loc[:,'firsttone']=tri_syl_data.loc[:,'tri_tones'].apply(lambda x:x.split(' ')[0])
            tri_syl_data.loc[:,'secondtone']=tri_syl_data.loc[:,'tri_tones'].apply(lambda x:x.split(' ')[1])
            tri_syl_data.loc[:,'thirdtone']=tri_syl_data.loc[:,'tri_tones'].apply(lambda x:x.split(' ')[2])
            tri_syl_table=tri_syl_data.groupby(['firsttone','secondtone','thirdtone'])['itemrec'].apply(lambda x:' // '.join(x))
            tri_syl_table=tri_syl_table.unstack(2)
        #要输出的结果表
        nowstr=datetime.now().strftime("%m%d")
        save_path=SOURCE.split('.xls')[0]+'分析(全部功能)结果'+nowstr+".xlsx"
        writer=pd.ExcelWriter(save_path)
        data.to_excel(writer,sheet_name='切分表',index=False,header=True)

        sv_cols=['项目','计数','','声母含例外','计数','声母','计数','韵母含例外','计数','韵母','计数','声调含例外','计数','声调','计数','韵尾含例外','计数','韵尾','计数']
        pd.DataFrame([['']*len(sv_cols) for i in range(2)],columns=sv_cols).to_excel(writer,index=False,sheet_name='概况')
        survey.to_excel(writer,header=False,sheet_name='概况',startrow=1,)
        init_list.to_excel(writer,header=False,sheet_name='概况',startrow=1,startcol=3)
        no_excep_init.to_excel(writer,header=False,sheet_name='概况',startrow=1,startcol=5)
        fnl_list.to_excel(writer,header=False,sheet_name='概况',startrow=1,startcol=7)
        no_excep_fnl.to_excel(writer,header=False,sheet_name='概况',startrow=1,startcol=9)
        tone_list.to_excel(writer,header=False,sheet_name='概况',startrow=1,startcol=11)
        no_excep_tone.to_excel(writer,header=False,sheet_name='概况',startrow=1,startcol=13)
        tail_list.to_excel(writer,header=False,sheet_name='概况',startrow=1,startcol=15)
        no_excep_tail.to_excel(writer,header=False,sheet_name='概况',startrow=1,startcol=17)
        if len(ned)>0:
            init_table_with_examples.to_excel(writer,sheet_name='音系')
            fnl_table_with_examples.to_excel(writer,sheet_name='音系',startrow=len(init_table_with_examples.index)+4)
            tone_phon.to_excel(writer,sheet_name='音系',startrow=len(init_table_with_examples.index)+len(fnl_table_with_examples)+8)

            homophone_all.to_excel(writer,sheet_name='同音字表')
            tone_match_init_fnl.to_excel(writer,sheet_name='音节配调表')
        if len(two_syl_data)>0:
            two_syl_table.to_excel(writer,sheet_name='声调组合表')
        if len(tri_syl_data)>0:
            tri_syl_table.to_excel(writer,sheet_name='声调组合表',startcol=len(two_syl_table)+4)
        
        #联系作者
        add='https://github.com/Lykit01/Lykit-for-chinese-dialects-field-work'
        about=pd.Series(['@Hue Zhang','3275803255@qq.com',add,'如有问题请联系我！'],index=['author','email','github',''])
        about.to_excel(writer,sheet_name='关于')
        
        writer.save()
        note_lb.config(text="分析完毕，结果保存在:"+save_path)
    #except Exception:
        #note_lb.config(text="表内数据有误！请检查源数据！")


# In[5]:


#选择源文件
def choose_files():
    #每次点击选择源文件，都会清空之前的选择,并且把wishlistname、data_ready等清空
    file_lb.config(text="")
    global TARGETSHEET
    global SOURCE
    global DATA_READY
    SOURCE=""
    DATA_READY=False
    
    filename=tkinter.filedialog.askopenfilenames()
    if len(filename)!=1:
        lb_text="你选择了%d个文件，请只选择1个文件。"%len(filename)
    elif filename[0].split('.')[-1] not in ['xls','xlsx','xlsm']:
        lb_text="请选择excel文件，后缀名包括xls,xlsx,xlsm等。"
    else:
        try:
            tryreadfile=pd.read_excel(filename[0],sheet_name=TARGETSHEET.get())
        except Exception:
            lb_text='选取文件不对或excel中没有唯一的名为'+TARGETSHEET.get()+'的sheet。'
        else:
            SOURCE=filename[0]
            DATA_READY=True
            lb_text="数据已经准备好了！"
            file_lb.config(text="您选择的文件是："+filename[0])
    note_lb.config(text=lb_text)


# In[ ]:





# In[12]:


root=Tk()
root.title("Lykit01:cdfw")
root.geometry("420x165")
SOURCE=''
DATA_READY=False
save_path=""

#提示分析中间的问题
note_lb=Label(root,text="请选择1个源文件,并在下面填入要处理的sheet名以及例外判断标准。")
note_lb.place(x=200,y=10,anchor="center")
#文件提示
file_lb=Label(root,text='')
file_lb.place(x=200,y=30,anchor="center")
#输入sheet名
Label(root,text='要处理的sheet名').place(x=50,y=60,anchor="center")
TARGETSHEET=StringVar()
TARGETSHEET.set('词表')
Entry(root,textvariable=TARGETSHEET,width=8).place(x=150,y=60,anchor="center")
Label(root,text='例字数目').place(x=230,y=60,anchor="center")
EXAMPLE_NUM=StringVar()
EXAMPLE_NUM.set(5)
Entry(root,textvariable=EXAMPLE_NUM,width=4).place(x=300,y=60,anchor="center")

Label(root,text='例外判定标准:声母').place(x=50,y=90,anchor="center")
INIT_EXCEP=StringVar()
INIT_EXCEP.set(0)
Entry(root,textvariable=INIT_EXCEP,width=4).place(x=150,y=90,anchor="center")
Label(root,text='韵母').place(x=200,y=90,anchor="center")
FNL_EXCEP=StringVar()
FNL_EXCEP.set(0)
Entry(root,textvariable=FNL_EXCEP,width=4).place(x=250,y=90,anchor="center")
Label(root,text='声调').place(x=300,y=90,anchor="center")
TONE_EXCEP=StringVar()
TONE_EXCEP.set(0)
Entry(root,textvariable=TONE_EXCEP,width=4).place(x=350,y=90,anchor="center")

choose_file_btn=Button(root,text="选择源文件",command=choose_files)
choose_file_btn.place(x=80,y=120,anchor="center")
execute_btn=Button(root,text="分析(全部功能)",command=execute)
execute_btn.place(x=180,y=120,anchor="center")
part_execute_btn=Button(root,text="分析(只切分声韵调)",command=part_execute)
part_execute_btn.place(x=300,y=120,anchor="center")
Label(root,text='@Hue Zhang 制作 如有问题请联系:3275803255@qq.com').place(x=200,y=150,anchor="center")

root.mainloop()


# In[ ]:




