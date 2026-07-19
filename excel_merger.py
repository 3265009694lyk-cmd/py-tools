# Excel表格合并工具


import pandas as pd
import os
import glob

def hebing(wenjianjia, shuchu, guanjianlie=None):
    # 文件夹在不在
    if not os.path.exists(wenjianjia):
        print('没找到这个文件夹')
        return

    # 找所有excel文件
    xlsx_list = glob.glob(os.path.join(wenjianjia, '*.xlsx'))
    xls_list = glob.glob(os.path.join(wenjianjia, '*.xls'))
    csv_list = glob.glob(os.path.join(wenjianjia, '*.csv'))

    all_files = xlsx_list + xls_list + csv_list

    if len(all_files) == 0:
        print('这个文件夹里没有表格文件')
        return

    print('找到了 %d 个文件' % len(all_files))
    print('-' * 30)

    # 存所有数据
    suoyou_shuju = []

    # 依次读取
    for f in all_files:
        try:
            name = os.path.basename(f)
            if f.endswith('.csv'):
                df = pd.read_csv(f, encoding='utf-8')
            else:
                df = pd.read_excel(f)

            # 标记来源
            df['来源文件'] = name

            suoyou_shuju.append(df)
            print('  OK: %s (%d行)' % (name, len(df)))
        except Exception as e:
            print('  读取失败: %s - %s' % (os.path.basename(f), str(e)))

    if len(suoyou_shuju) == 0:
        print('没读到任何数据')
        return

    # 合并
    result = pd.concat(suoyou_shuju, ignore_index=True)
    print('合并之后共 %d 行' % len(result))

    # 开始清洗

    # 1. 去掉列名两边的空格
    cols = []
    for c in result.columns:
        cols.append(c.strip())
    result.columns = cols

    # 2. 去掉完全重复的行
    qian = len(result)
    result = result.drop_duplicates()
    qudiao = qian - len(result)
    if qudiao > 0:
        print('去掉了 %d 行重复的' % qudiao)

    # 3. 如果有关键列，按关键列再去一次重
    if guanjianlie and guanjianlie in result.columns:
        qian = len(result)
        result = result.drop_duplicates(subset=[guanjianlie], keep='first')
        you_qudiao = qian - len(result)
        if you_qudiao > 0:
            print('按[%s]去重: 去掉了 %d 行' % (guanjianlie, you_qudiao))

    # 4. 处理空值
    for lie in result.columns:
        kong = result[lie].isnull().sum()
        if kong > 0:
            print('%s 有 %d 个空值，已填为无' % (lie, kong))
            result[lie] = result[lie].fillna('无')

    # 5. 输出
    if shuchu.endswith('.csv'):
        result.to_csv(shuchu, index=False, encoding='utf-8-sig')
    else:
        if not shuchu.endswith('.xlsx'):
            shuchu = shuchu + '.xlsx'
        result.to_excel(shuchu, index=False)

    print('完成，最后是 %d 行数据' % len(result))
    print('保存到了: %s' % shuchu)
    return result

if __name__ == '__main__':
    import sys
    print('===== Excel合并清洗工具 =====')
    if len(sys.argv) >= 3:
        folder = sys.argv[1]
        out = sys.argv[2]
        key = sys.argv[3] if len(sys.argv) > 3 else None
    else:
        folder = input('表格文件所在的文件夹: ')
        out = input('输出文件名(例如 结果.xlsx): ')
        k = input('去重关键列(例如 学号，直接回车跳过): ')
        key = k if k.strip() != '' else None
    hebing(folder, out, key)
    input('按回车退出...')
