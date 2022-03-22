# 処理一覧
import re


class GetRightNumber:
    def __init__(self):
        self.name = '右側から数字を抜き出す'
        self.explain = 'サンプルサンプルサンプルサンプル'

    def run(self, src_ws, tgt_ws, tgt):
        c = 1  # col 通し番号
        matcher = re.compile('\d+$')  # 数字
        for cols in tgt:
            for col in cols:
                r = 1  # row　通し番号
                target_col = src_ws[col]
                for i in range(len(target_col)):
                    cell = src_ws[col][i].value  # col列のi行目
                    if cell is not None:
                        result = re.search(matcher, cell.rstrip())
                        if result:
                            # matchした場合
                            tgt_ws.cell(row=r, column=c).value = cell[:result.start()]
                            tgt_ws.cell(row=r, column=c + 1).value = int(result.group())
                        else:
                            # matchしない場合そのまま書き込み
                            tgt_ws.cell(row=r, column=c).value = cell
                    r += 1
                c += 2
            c += 1


class Sample2:
    def __init__(self):
        self.name = 'サンプル2'
        self.explain = 'サンプル2サンプル2サンプル2サンプル2'

    def run(self):
        print('sample2')
        state = '成功'
        return state
