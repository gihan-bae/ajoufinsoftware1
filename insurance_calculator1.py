import openpyxl

MAN = 0     # 남자 코드 = 0
WOMAN = 1   # 여자 코드 = 0

INSURANCE_TYPE0 = 0 # 정기보험
INSURANCE_TYPE1 = 1 # 종신보험
INSURANCE_TYPE2 = 2 # 연금보험
# 보험료 계산 클래스 #
class InsuruanceCalculator:
    # 컬럼명 리스트 #
    columns = ['year', 'live_all', 'live_man', 'live_woman',
               'dead_all', 'dead_man', 'dead_woman']

    # 생성자 #
    def __init__(self):
        self.data_list = self.load_data() # 변수 self.data_list에 self.load_data() 메소드의 반환값 저장

    # 데이터 불러오기 #
    def load_data(self):
        workbook = openpyxl.load_workbook('2019생명표.xlsx') # openpyxl 모듈을 사용해 '2019생명표.xlsx' 파일을 연다

        sheet = workbook['데이터'] # '데이터' 시트를 변수 sheet에 저장
        data_list = [] # 데이터 저장할 리스트 생성
        # ord() 함수는 인자로 주어진 문자의 아스키 코드 값을 반환해 줌
        # 따라서 리스트 cell_map에 ['A','H','I','J','Q','R','S']의 각각 요소들을 참조하여
        # 각 요소들의 아스키 코드 값 - A의 아스키 코드값을 저장
        cell_map = [ord(x)-ord('A') for x in ['A','H','I','J','Q','R','S']]
        # enumerate() 함수는 인자로 주어진 연속 가능한 객체(list, tuple, str 등)의 각 요소를 (인덱스, 값)의 형태로 반환해줌
        # '데이터' 시트의 각 행들을 enumerate() 함수를 사용해 i에 인덱스, row에 행을 저장하며 반복
        for i, row in enumerate(sheet.rows):
            if i == 0:   # 인덱스가 0일 때
                continue # 다음 반복으로 넘어감
            data = {} # 데이터 딕셔너리 생성
            # 클래스 내에 생성한 컬럼명 리스트를 enumerate() 함수를 사용해 j에 인덱스, column에 컬럼명을 저장하며 반복
            for j, column in enumerate(InsuruanceCalculator.columns):
                idx = cell_map[j]               # idx에 리스트 cell_map의 j번 인덱스 값 저장
                data[column] = row[idx].value   # 데이터 딕셔너리에 column명: 행의 idx번째 컬럼의 값 저장
            data_list.append(data) # 리스트 data_list에 딕셔너리 data 추가
        return data_list # 리스트 data_list 반환

    def v(self, i):
        return 1/(i+1)

    # x : 나이, i : 이자율
    def C(self, x, i, sex):
        column = InsuruanceCalculator.columns[5 + sex] # column에 컬럼명 리스트의 5+sex번 인덱스 값 저장
        dx = self.data_list[x][column] # dx에 data_list의 x번 인덱스의 딕셔너리 데이터 중 키가 column인 값 저장
        return dx * pow(self.v(i), x+1) # pow() 함수는 제곱을 해주는 함수, 즉, dx * (v(i))**(x+1)의 계산 결과 값을 반환함
    def M(self, x, i, sex):
        result = 0
        for x in range(x, 101): # x부터 100까지 반복
            result += self.C(x, i, sex) # result에 C(x, i, sex)의 반환값을 더해줌
        return result # result 반환
    def D(self, x, i, sex):
        column = InsuruanceCalculator.columns[2 + sex] # column에 컬럼명 리스트의 2+sex번 인덱스 값 저장
        lx = self.data_list[x][column] # lx에 data_list의 x번 인덱스의 딕셔너리 데이터 중 키가 column인 값 저장
        return lx * pow(self.v(i), x) # lx * (v(i))**x의 결과값 반환
    def N(self, x, i, sex):
        result = 0
        for x in range(x, 101): # x부터 100까지 반복
            result += self.D(x, i, sex) # result에 D(x, i, sex)의 반환값을 더해줌
        return result # result 반환

    def P(self, x, i, sex, n):
        m = self.M(x, i, sex)
        m_n = self.M(x + n, i, sex)
        n_val = self.N(x, i, sex)
        n_n_val = self.N(x + n, i, sex)
        return (m-m_n)/(n_val-n_n_val) # => (M(x)-M(x+n)) / (N(x)-N(x+n))

    def NSP(self, x, i, sex, n, A, m_year, insur_type):
        if insur_type == INSURANCE_TYPE0: # 정기보험일 때
            m = self.M(x + m_year, i, sex)
            m_n = self.M(x + n + m_year, i, sex)
            d = self.D(x, i, sex)
            return ((m-m_n)/d) * A # => [(M(x)-M(x+n))/D(x)] * A
        elif insur_type == INSURANCE_TYPE1: # 종신보험일 때
            m = self.M(x+m_year, i, sex)
            d = self.D(x, i, sex)
            return (m/d) * A # (m / d) * A 반환 => (M(x)/D(x)) * A
        else: # 연금보험일 때
            n_val = self.N(x + m_year + 1, i, sex)
            n_n_val = self.N(x + m_year + n + 1, i, sex)
            d = self.D(x, i, sex)
            return ((n_val - n_n_val) / d) * A # => [(N(x)-N(x+n))/D(x)] * A

    def NMP(self, x, i, sex, n, A, m_year, insur_type):
        if insur_type == INSURANCE_TYPE0: # 정기보험일 때
            m = self.M(x + m_year, i, sex)
            m_n = self.M(x + n + m_year, i, sex)
            n_val = self.N(x, i, sex)
            n_n_val = self.N(x + n, i, sex)
            d = self.D(x, i, sex)
            d_n = self.D(x + n, i, sex)
            return (m-m_n)/((n_val-n_n_val) - (11/24)*(d-d_n)) * A * (1/12) # => (M(x)-M(x+n))/[(N(x)-N(x+n))-(11/24)*(D(x)-D(x+n))] * A * (1/12)
        elif insur_type == INSURANCE_TYPE1: # 종신보험일 때
            m = self.M(x + m_year, i, sex)
            n_val = self.N(x, i, sex)
            d = self.D(x, i, sex)
            return (m)/((n_val) - (11/24)*(d)) * A * (1/12) # => M(x) / [N(x)-(11/24)*D(x)] * A * (1/12)


if __name__ == '__main__':
    ic = InsuruanceCalculator() # InsuruanceCalculator 클래스 객체 생성
    type_map = {t:i for i,t in enumerate(['정기보험', '종신보험', '연금'])} # type_map을 키가 보험 종류, 값이 인덱스인 딕셔너리로 생성
    # 무한반복
    while True:
        insur_type = input("보험 종류를 입력해주세요 (정기보험, 종신보험, 연금) : ") # 보험 종류 입력
        if insur_type not in type_map: # 입력받은 보험이 딕셔너리 type_map에 존재하지 않을 때
            print("보험 종류를 다시 입력해주세요\n\n")
            continue # 다음 반복으로 넘어감
        insur_type = type_map[insur_type] # 딕셔너리 type_map에 입력받은 보험 종류에 해당하는 값을 insur_type에 저장
        print(insur_type) # 출력
        sex = input("성별 : (MAN/WOMAN) ").lower() # 성별을 MAN, 혹은 WOMAN으로 입력받고, 소문자로 변경하여 sex에 저장(대소문자 구분 없이 확인하기 위한 용도)
        sex = MAN if sex == 'man' else WOMAN# 성별이 남자일 땐 sex에 0, 여자일 땐 1 저장
        print("나이와 거치년도 가입기간의 합은 100을 초과할 수 없습니다")
        year = int(input("나이 : ")) # 나이 입력
        year = year if year < 100 else 100 # 나이가 100살 미만일 땐 그대로 저장하고, 100세 이상일 땐 100세로 저장
        m_year = int(input("거치년도 : ")) # 보험 거치 년도 입력
        if insur_type != INSURANCE_TYPE1: # 종신보험이 아닐 때
            date = int(input("가입 기간 : ")) # 가입 기간 입력받음
        else: # 종신보험일 때
            date = -1 # date에 -1 저장

        A = float(input("받을 보험금 : ")) # 받을 보험금 입력
        i = float(input("이자율 : ")) # 이자율 입력, 5%면 0.05

        nsp = ic.NSP(year, i, sex, date, A, m_year, insur_type) # NSP 계산
        if insur_type != INSURANCE_TYPE2: # 연금보험이 아닐 때
            nmp = ic.NMP(year, i, sex, date, A, m_year, insur_type) # NMP 계산

        print("일시납순보험료 : %f" % nsp) # NSP 출력
        if insur_type != INSURANCE_TYPE2: # 연금보험이 아닐 때
            print("월납보험료 : %f" % nmp) # NMP 출력
        print("\n\n")
        inpt = input("더 계산하시겠습니까? (Y/N)").lower()
        if inpt == 'n':
            break