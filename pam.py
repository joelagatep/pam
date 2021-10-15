import pandas as pd

def main():
    df = pd.read_excel('PAMPSG-20211005.xls')
    for i in range(len(df)):
        members_wd_acct = str(df.iloc[i,6])
        user_based = df.iloc[i,9]
        #print(member + "," + group + "," + members_wd_acct + "," + str(members_wd_acct.find("wd-")) + "," + str(members_wd_acct.find("-impl")))
        if user_based == "NO" \
        and members_wd_acct.find("wd-",0,3) == -1 \
        and members_wd_acct.find("-impl") == -1 :
            group = df.iloc[i,0]
            member = df.iloc[i,1]
            print(f"{member},{group},{members_wd_acct}")


if __name__ == '__main__':
    main()



