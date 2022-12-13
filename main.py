from sys import exit
from os import system
from time import sleep
from pandas import read_excel
from math import sqrt

def intro():
    introduction_text = """
    ------------------------------------------------------
                    Tolerance Calculator
                                            - Young
    ------------------------------------------------------

    data.xlsx에 치수 누적 방향, 치수, 공차를 입력한 후 실행
    """
    print(introduction_text)

    while(True):
        answer = input("계속 진행?(y/n) : ").lower()
        if answer == "n":
            exit(0)
        elif answer == "y":
            break;

def main():
    data = read_excel("./data.xlsx", header=0)
    headers = data.columns
    if '방향' not in headers and '치수' not in headers and '공차' not in headers:
        print("올바른 data.xlsx 파일을 사용하여 값을 입력한 후 다시 실행해 주세요.")
        return
    direction = data.loc[:, '방향']
    length = data.loc[:, '치수']
    tolerance = data.loc[:, '공차']

    if not (direction.count() == length.count() and length.count() == tolerance.count()):
        print("data.xlsx 파일에 방향, 치수, 공차값이 모두 입력했는지 확인 후 다시 실행해 주세요.")
        return

    total_count = direction.count()
    if not total_count > 1:
        print("data.xlsx 파일에 최소 2개 이상의 데이터를 입력한 후 다시 실행해 주세요.")
        return
    
    total_length = 0.0
    for i in range(0, total_count):
        if direction.loc[i] == "-":
            total_length -= length.loc[i]
        else:
            total_length += length.loc[i]

    sum_total, sum_square = 0.0, 0.0
    for i in range(total_count):
        sum_total += tolerance.loc[i]
        sum_square += tolerance.loc[i]**2
    root_sum_square = sqrt(sum_square)

    total_tolerance_wc = round(sum_total, 3)
    total_tolerance_rss = round(root_sum_square, 3)
    total_tolerance_comb = round( ((total_count - 2)*total_tolerance_rss + 2*total_tolerance_wc)/total_count, 3 )

    print(f"WC: {round(total_length,3)} ± {total_tolerance_wc}")
    print(f"RSS: {round(total_length,3)} ± {total_tolerance_rss}")
    print(f"Combination: {round(total_length,3)} ± {total_tolerance_comb}")

if __name__ == "__main__":
    system('cls')
    intro()

    while(True):
        system('cls')
        sleep(1)
        main()
        # print(value_list, tolerance_list)
        sleep(1)
        while(True):
            answer = input("다시 계산하기 (y/n) : ").lower()
            if answer == "n":
                exit(0)
            elif answer == "y":
                break;
