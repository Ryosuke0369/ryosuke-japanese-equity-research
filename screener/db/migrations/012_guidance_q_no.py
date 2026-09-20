"""guidance に q_no（予想の対象期間）を足す。

業績予想の修正開示には「中間期だけの修正」がある（3161 は H1 のみ）。
guidance のキーは (code, date, fy, item) しか無いので、通期予想と中間期予想が
同じ行に載りうる。通期を優先して保存しつつ、**どちらを保存したのか**を
読み手が判別できるように列に残す。S5 の進捗率は通期予想を分母にする。
"""


def up(con):
    cols = [r[1] for r in con.execute("PRAGMA table_info(guidance)")]
    if "q_no" not in cols:
        con.execute("ALTER TABLE guidance ADD COLUMN q_no INTEGER")
