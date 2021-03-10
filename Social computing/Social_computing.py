#将所有评论用工具分词
#选定正向负向情感种子词各5个
#利用公式计算每个单词的语义倾向值：t = ∑ PMI( word , pword ) - ∑ PMI( word , nword )
#其中，PMI( word1 , word2 ) = log2( ( P( word1 ∩ word2 ) / ( P( word1 ) * P( word2 ) ) ) )
#倾向值最大的50个和最小的50个单词即为所求

import xlrd
import jieba
import re,string
import math

def get_words():
   #将所有评论分词后存入word_list列表    
    punc = '~`!#$%^&*()_+-=|\';":/.,?><~·！@#￥%……&*（）——+-=“：’；、。，？》《{}～, '
    file_path = 'D:\File_Yunshou\SN赛后采访视频评论.xlsx'
    word_list = [ ]
   
    data = xlrd.open_workbook( file_path )
    table = data.sheets( )[ 0 ]
    cols_value = table.col_values( 0 )
    cols_value = ", ".join( cols_value )
    cols_value = re.sub( r"[%s]+" %punc , "" , cols_value ) 
    allwords = jieba.lcut( cols_value )
    word_list = list( set( allwords ) )
    return word_list

def get_sentence():
    #将所有评论语句收集存入sent_list列表    
    file_path = 'D:\File_Yunshou\SN赛后采访视频评论.xlsx'
    punc = '~`!#$%^&*()_+-=|\';":/.,?><~·！@#￥%……&*（）——+-=“：’；、。，？》《{}～, '
    sent_list = [ ]
    data = xlrd.open_workbook( file_path )
    table = data.sheets( )[ 0 ]
    sent_list = table.col_values( 0 )
    return sent_list

def p_calcu( word_list , sent_list ):
    p_word_list = { }
    length = len( sent_list )
    #计算各个分词出现的频率
    #将计算后的频率存入p_word_list字典
    for word in word_list :
        count = 0
        for sent in sent_list :
            if sent.find( word ) != -1 :
                count += 1
        p_word_list[ word ] = count / length
    return p_word_list
 
def pp_calcu( word_list , seeds , sent_list):
    pp_word_list = { }
    length = len( sent_list )
    #计算各个单词和基准词的联合出现频率  
    #计算每个单词的语义倾向值
    for word in word_list :
        count = 0
        pp_word_list[ word ] = { }
        for seed in seeds :
            for sent in sent_list :
                if (sent.find( word ) != -1) and (sent.find( seed ) != -1):
                    count +=1
                pp_word_list[ word ][ seed ] = count / length
    return pp_word_list

def pmi_sort( p_word_list , p_seed_p , p_seed_n , pp_word_p, pp_word_n ):
    pmi_so = { }
    for word in p_word_list.keys() :
        pmi_p = 0
        pmi_n = 0
        for p in p_seed_p.keys() :
            pmi_p += math.log( ( pp_word_p[ word ][ p ] +1 ) / ( p_seed_p[ p ] * p_word_list[ word ] + 1 ) , 2 )
        for n in p_seed_n.keys() :
            pmi_n += math.log( ( pp_word_n[ word ][ n ] +1 ) / ( p_seed_n[ n ] * p_word_list[ word ] + 1 ) , 2 )
        pmi_so[ word ] = (pmi_p - pmi_n)*10
    pmi_so = sorted( pmi_so.items() , key=lambda x : x[ 1 ] , reverse = False )
    return pmi_so

seeds = [ '加油' , '未来可期' , '精彩' , '血性' , '感谢', '尽力', '卷土重来' , '谢谢', '不错', '稳健','失望' , '难受' , '内战幻神' , '拉胯' , '遗憾', '舒服了', '笑死', '梦游', '解散', '菜' ]
seed_p = [ '加油' , '未来可期' , '精彩' , '血性' , '感谢', '尽力', '卷土重来' , '谢谢', '不错', '稳健']
seed_n = [ '失望' , '难受' , '内战幻神' , '拉胯' , '遗憾', '舒服了', '笑死', '梦游', '解散', '菜' ]
word_list = get_words()
print('word_list列表已取得！')
sent_list = get_sentence()
print('sent_list列表已取得！')
p_word_list = p_calcu( word_list , sent_list )
print('p_word_list字典已取得！')
p_seed_p = p_calcu( seed_p , sent_list )
print('p_seed_p字典已取得！')
p_seed_n = p_calcu( seed_n , sent_list )
print('p_seed_n字典已取得！')
pp_word_p = pp_calcu( word_list , seed_p , sent_list)
print('pp_word_p字典已取得！')
pp_word_n = pp_calcu( word_list , seed_n , sent_list)
print('pp_word_n字典已取得！')
pmi_so = pmi_sort( p_word_list , p_seed_p , p_seed_n , pp_word_p, pp_word_n )
with open('D:\File_Yunshou\pmi_so排序结果.txt', 'w', encoding='utf-8') as f:
    for s,obj in pmi_so:
        f.write(str(obj))
        f.write("\n")
print("写入完成！")
