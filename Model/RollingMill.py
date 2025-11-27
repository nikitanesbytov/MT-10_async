from math import *

class RollingMill:
    def __init__(self,DR,L,b,h_0,StartTemp,DV,MV,MS,OutTemp,SteelGrade,V0,V1,S,V_Valk_Per,StartS,d1,d2,d,VS,Dir_of_rot):
        #Параметры сляба(Задает оператор)
        self.L = L #Начальная длина сляба(мм)
        self.b = b #Ширина сляба(мм)
        self.h_0 = h_0 #Начальная толщина сляба(мм)
        self.h_1 = S #Конечная толщина сляба(мм)
        self.StartTemp = StartTemp #Начальная температура сляба(Температура выдачи из печи)(°C)
       
        #Параметры по умолчанию
        self.ZK = 0 #Жесткость клети
        self.DV = DV #Диаметр валков(мм)
        self.DR = DR #Диаметр рольгангов(мм)
        self.R = DV/2 #Радиус валков(мм)
        self.MV = MV #Материал валков
        self.MS = MS  #Материал сляба
        self.TempV = OutTemp #Температура валков(°C)
        self.d1 = d1 #Расстояние пути до валков(мм)
        self.d2 = d2 #Расстояние пути после валков(мм)
        self.d = d #Расстояние между левыми и правыми рольгангами

        #Настройка ТП
        self.CurrentS = StartS #Нынешнее положение раствора валков
        self.V1 = V1 #Скорость рольгангов после валков(мм/c)
        self.V0 = V0 #Скорость рольгангов до валков(мм/c)
        self.VS = VS #Скорость выставления валков(мм/c)
        self.accel = 500 #Рагон валков и рольгангов(мм/c2)
        self.stop_accel = 500 #Ускорение замедления(мм/c2)
        self.S = S #Раствор валков(Массив)(Задает оператор)(мм)
        self.V_Valk_Per = V_Valk_Per #Заданная скорость валков оператором в об/мин
        self.SteelGrade = SteelGrade #Марка стали
        self.Dir_of_rot = Dir_of_rot #Направление варщения 
 

    def SpeedOfRolling(self, DV, V) -> float:
        # w = V / (π * DV) - частота вращения (об/с)
        # где V в мм/с, DV в мм
        w = V / (pi * DV)  
        # Скорость прокатки = π * DV * w * (1 + 0.05)
        SpeedOfRolling = pi * DV * w * (1 + 0.05)
        return SpeedOfRolling/1000 #мм/с
    
    def RelDef(self, h_0, h_1) -> float:
        "Относительная деформация"
        RelDef = (h_0 - h_1) / h_0
        return RelDef
    
    def TempDrDConRoll(self, DV, h_0, h_1, Temp, SpeedOfRolling) -> float:
        "Падение температуры вследствие контакта с валками"
        # Все размеры в мм, скорость в мм/с
        delta_h = h_0 - h_1  # абсолютное обжатие (мм)
        
        # Вычисление угла контакта (в радианах)
        contact_angle = acos(1 - delta_h / DV)
        
        # Длина контакта (мм)
        contact_length = sqrt((DV/2) * contact_angle)
        
        TempDrDConRoll = (0.216 * contact_length / (h_0 + h_1) * 
                        (Temp - 60) * sqrt(1.08 / SpeedOfRolling) * 0.8)
        
        return TempDrDConRoll  # в °C (падение температуры)
   
    def TempDrPlDeform(self,DefResistance,h_0,h_1) -> float:
        "Прирост температуры вследствие пластической деформации"
        TempDrPlDeform = 0.183 * (DefResistance) * log(h_0/h_1)
        #RelDef - Степень деформации
        #DefResistance - Сопротивление деформации(МПа) 
        return TempDrPlDeform
    
    def GenTemp(self,Temp,TempDrBPass,TempDrDConRoll,TempDrPlDeform) -> float:
        "Общая температура после итерации прокатки"
        GenTemp = Temp + TempDrPlDeform - TempDrDConRoll - TempDrBPass
        #Temp - Температура в проходе(°C)
        return GenTemp

    def DefResistance(self,RelDef,LK,V,CurrentTemp,SteelGrade) -> float:
        "Сопротивление деформации(МПа)"
        SteelGrades = {"Ст3сп":(87.1,0.124,0.167,2.8),
                      "12ХН3А":(89.9,0.095,0.261,2.84),
                      "65Г":(73.2,0.166,0.222,3.02),
                      "К65":(83.2,0.149,0.213,4.143),
                      "X100":(84.5,0.161,0.197,4.208),
                      "HARDOX500":(92.3,0.159,0.291,3.756),
                      "08Х18Н10Т":(175.4,0.1312,0.1493,4.2269)}
        sigmaOD,a,b,c = SteelGrades[SteelGrade] 
        u = (V/LK * RelDef)
        Sigmaf = sigmaOD * (u**a) * ((10*RelDef)**b) * ((CurrentTemp/1000)**-c)
        #a,b,c - Коэффициенты зависящие от марки стали
        #u - Средняя скорость деформации(1/c)
        #V - Скорость валков(мм/c)
        #sigmaOD - Базисное значение сопротивления деформации
        #RelDef - Степень деформации
        #CurrentTemp - Нынешняя температура
        #LK - Длина дуги контакта(мм)
        return Sigmaf #[Мпа]
    
    def Moment(self,LK,h_0,h_1,Effort):
        "Расчет момента прокатки(кНм)"
        h_average = (h_1 + h_0)/2
        psi = 0.498 - 0.283 * LK / h_average
        Moment = 2 * Effort * psi * LK
        return Moment
    
    def Effort(self,LK,b,AvrgPressure):
        "Расчет усилия прокатки(Н)"
        F = LK * b
        P = AvrgPressure * F
        return P
    
    def Power(self,M,V,DV) -> float:
        "Рассчет мощности прокатки(Вт)"
        w = V / (pi * DV) 
        N = M * w 
        # М - Крутящий момент на валках(Н*м)
        # omega - Угловая скорость вращения валков(рад/c)
        # N - Мощность прокатки(Вт)
        return N

    def CapCondition(self, Mu, S, DV) -> bool:
        "Условие захвата"
        alpha = acos(1 - (S / DV))
        CapCon = (Mu >= tan(alpha))
        return CapCon
    
    def FricCoef(self,MV, MS, V0, TempS) -> float:
        "Коэффициент трения"
        if (MV == 'Сталь'): 
            k1 = 1
        elif(MV == 'Чугун'):
            k1 = 0.8

        if (V0 <= 3):
            k2 = 0.8
        elif (V0 > 3):
            k2 = 1.53 * V0 ** (-0.47)

        if (MS == 'Carbon Steel'): 
            k3 = 1
        elif (MS == 'Austenitic steel'):
            k3 = 1.47

        Mu = k1 * k2 * k3 * (1.05 - 0.0005 * TempS) 
        return Mu
    
    def AvrgPressure(self,LK,h_1,h_0,DefResistance) -> float:
        "Среднее давление на валки"
        h_average = (h_1 + h_0)/2
        if ((LK/h_average) <= 2):
            n_frict = 1 + (LK/h_average)/6
        elif (((LK/h_average) > 2) and ((LK/h_average) <= 4)):
            n_frict = 1 + (LK/h_average)/5
        elif ((LK/h_average) > 4):
            n_frict = 1 + (LK/h_average)/4

        if LK/h_average < 1:
            n_zone = (LK / h_average) ** -0.4
        else:
            n_zone = 1
       
        P = 1.15 * n_frict * n_zone * DefResistance 
        return P #[Мпа]
    
    def ContactArcLen(self,DV,h_0,h_1) -> float:
        "Длина дуги контакта"
        # LK = sqrt(DV/2 * (h_0 - h_1))
        delta_h = h_0 - h_1
        LK = sqrt((DV * delta_h)/2)
        return LK #мм
    
    def TempDrBPass(self, T0, Time,width,height):
        "Падение температуры между пропусками"
        P0 =  2 * (width + height) # Периметр поперечного сечения
        F0 = width * height  # Площадь поперечного сечения
        cube_root = (0.0255 * P0 * Time / F0 + (1000 / (T0 + 273)) ** 3) ** (1/3)
        delta_t0 = T0 - (1000 / cube_root) + 273
        return delta_t0
    
    # def ContactArea(self,b_0, b_1, LK) -> float:
    #     "Площадь контакта"
    #     b_average = (b_0 + b_1) / 2
    #     F = b_average * LK
    #     return F



     

