import sys,os
from collections import OrderedDict
from datetime import datetime
import locale
locale.setlocale(locale.LC_ALL, '')


commandDict = OrderedDict()

def abc(i):
    return chr(ord('@')+i+1)

PIList = sys.argv[1].split(",")

# PIList = ["SS","TH","LL"] # EF
# PIList += ["YE"] # Total


sums = {}

for item in [
    "SalaryAndBenefitsTotalYearOne",
    "SalaryAndBenefitsTotalYearTwo",
    "SalaryAndBenefitsTotalYearThree",
    "SalaryAndBenefitsTotalYearFour",
    "SalaryAndBenefitsTotalYearFive",
    "SalaryAndBenefitsTotal",
    "TravelTotalYearOne",
    "TravelTotalYearTwo",
    "TravelTotalYearThree",
    "TravelTotalYearFour",
    "TravelTotalYearFive",
    "TravelTotal",
    "TuitionYearOne",
    "TuitionYearTwo",
    "TuitionYearThree",
    "TuitionYearFour",
    "TuitionYearFive",
    "TuitionTotal",
    "OtherDirectTotalYearOne",
    "OtherDirectTotalYearTwo",
    "OtherDirectTotalYearThree",
    "OtherDirectTotalYearFour",
    "OtherDirectTotalYearFive",
    "OtherDirectTotal",
    "IndirectTotalYearOne",
    "IndirectTotalYearTwo",
    "IndirectTotalYearThree",
    "IndirectTotalYearFour",
    "IndirectTotalYearFive",
    "IndirectTotal",
    "TotalYearOne",
    "TotalYearTwo",
    "TotalYearThree",
    "TotalYearFour",
    "TotalYearFive",
    "Total",
    ]:
    sums[item] = 0



def lines_that_contain(string, lines):
    return [line for line in lines if string in line]

def grabLineAndConvertToFloat(string,lines):
    tmpInput = lines_that_contain(string,lines)
    print(tmpInput)
    tmpInput = tmpInput[0]
    tmpInput = tmpInput.split("{")[-1]
    tmpInput = tmpInput.split("}")[0]
    try:
        tmpInput = locale.atof(tmpInput)
    except:
        tmpInput = 0
    return tmpInput
    # print(tmpInput)

for prefix in PIList: 

    # Open tex files and read in yearly totals
    # Add them and define new commands:
    # 
    f = open(prefix+"BudgetDefs.tex", "r")
    lines = f.readlines()
    sums["SalaryAndBenefitsTotalYearOne"] += grabLineAndConvertToFloat(prefix+"SalaryAndBenefitsTotalYearOne}",lines)
    sums["SalaryAndBenefitsTotalYearTwo"] += grabLineAndConvertToFloat(prefix+"SalaryAndBenefitsTotalYearTwo}",lines)
    sums["SalaryAndBenefitsTotalYearThree"] += grabLineAndConvertToFloat(prefix+"SalaryAndBenefitsTotalYearThree}",lines)
    sums["SalaryAndBenefitsTotalYearFour"] += grabLineAndConvertToFloat(prefix+"SalaryAndBenefitsTotalYearFour}",lines)
    sums["SalaryAndBenefitsTotalYearFive"] += grabLineAndConvertToFloat(prefix+"SalaryAndBenefitsTotalYearFive}",lines)
    sums["SalaryAndBenefitsTotal"] += grabLineAndConvertToFloat(prefix+"SalaryAndBenefitsTotal}",lines)


    sums["TravelTotalYearOne"] += grabLineAndConvertToFloat(prefix+"TravelTotalYearOne}",lines)
    sums["TravelTotalYearTwo"] += grabLineAndConvertToFloat(prefix+"TravelTotalYearTwo}",lines)
    sums["TravelTotalYearThree"] += grabLineAndConvertToFloat(prefix+"TravelTotalYearThree}",lines)
    sums["TravelTotalYearFour"] += grabLineAndConvertToFloat(prefix+"TravelTotalYearFour}",lines)
    sums["TravelTotalYearFive"] += grabLineAndConvertToFloat(prefix+"TravelTotalYearFive}",lines)
    sums["TravelTotal"] += grabLineAndConvertToFloat(prefix+"TravelTotal}",lines)

    sums["TuitionYearOne"] += grabLineAndConvertToFloat(prefix+"TuitionYearOne}",lines)
    sums["TuitionYearTwo"] += grabLineAndConvertToFloat(prefix+"TuitionYearTwo}",lines)
    sums["TuitionYearThree"] += grabLineAndConvertToFloat(prefix+"TuitionYearThree}",lines)
    sums["TuitionYearFour"] += grabLineAndConvertToFloat(prefix+"TuitionYearFour}",lines)
    sums["TuitionYearFive"] += grabLineAndConvertToFloat(prefix+"TuitionYearFive}",lines)
    sums["TuitionTotal"] += grabLineAndConvertToFloat(prefix+"TuitionTotal}",lines)


    sums["OtherDirectTotalYearOne"] += grabLineAndConvertToFloat(prefix+"OtherDirectTotalYearOne}",lines)
    sums["OtherDirectTotalYearTwo"] += grabLineAndConvertToFloat(prefix+"OtherDirectTotalYearTwo}",lines)
    sums["OtherDirectTotalYearThree"] += grabLineAndConvertToFloat(prefix+"OtherDirectTotalYearThree}",lines)
    sums["OtherDirectTotalYearFour"] += grabLineAndConvertToFloat(prefix+"OtherDirectTotalYearFour}",lines)
    sums["OtherDirectTotalYearFive"] += grabLineAndConvertToFloat(prefix+"OtherDirectTotalYearFive}",lines)
    sums["OtherDirectTotal"] += grabLineAndConvertToFloat(prefix+"OtherDirectTotal}",lines)

    sums["IndirectTotalYearOne"] += grabLineAndConvertToFloat(prefix+"IndirectTotalYearOne}",lines)
    sums["IndirectTotalYearTwo"] += grabLineAndConvertToFloat(prefix+"IndirectTotalYearTwo}",lines)
    sums["IndirectTotalYearThree"] += grabLineAndConvertToFloat(prefix+"IndirectTotalYearThree}",lines)
    sums["IndirectTotalYearFour"] += grabLineAndConvertToFloat(prefix+"IndirectTotalYearFour}",lines)
    sums["IndirectTotalYearFive"] += grabLineAndConvertToFloat(prefix+"IndirectTotalYearFive}",lines)
    sums["IndirectTotal"] += grabLineAndConvertToFloat(prefix+"IndirectTotal}",lines)

    sums["TotalYearOne"] += grabLineAndConvertToFloat(prefix+"TotalYearOne}",lines)
    sums["TotalYearTwo"] += grabLineAndConvertToFloat(prefix+"TotalYearTwo}",lines)
    sums["TotalYearThree"] += grabLineAndConvertToFloat(prefix+"TotalYearThree}",lines)
    sums["TotalYearFour"] += grabLineAndConvertToFloat(prefix+"TotalYearFour}",lines)
    sums["TotalYearFive"] += grabLineAndConvertToFloat(prefix+"TotalYearFive}",lines)
    sums["Total"] += grabLineAndConvertToFloat(prefix+"Total}",lines)
    # pass

print(sums)

combinedPrefix = "".join(PIList)

for key in sums:
    commandDict[key] = sums[key]

# commandDict[f"Total"] = TotalSum
# commandDict[f"TotalYearOne"] = TotalYearOneSum
# commandDict[f"TotalYearTwo"] = TotalYearTwoSum
# commandDict[f"TotalYearThree"] = TotalYearThreeSum
# commandDict[f"TotalYearFour"] = TotalYearFourSum
# commandDict[f"TotalYearFive"] = TotalYearFiveSum


with open(combinedPrefix+'BudgetDefs.tex', 'w') as outputFile:

    outputFile.write(f'% {datetime.now()}' + os.linesep)
    outputFile.write(f'% Auto-generated by {sys.argv[0]}' + os.linesep)
    outputFile.write(f'% >>> {" ".join(sys.argv)}' + os.linesep)
    outputFile.write(f'% >>> Written by L Lee for UTK Budget Spreadsheet as of 2022' + os.linesep)
    outputFile.write(os.linesep)


    for key,val in commandDict.items():
        if not val:
            val = ""
        try:
            # if math.modf(float(val))[0]:
            #     val = "%.2f" % val
            if float(val):
                # val = "%.2f" % val
                val = '{:,.2f}'.format(val)
                # val = "0"
        except:
            pass
        
        print("{: >30} {: >20}".format(key,val))

        outputFile.write("\\newcommand{\\%s}{%s}"%(combinedPrefix+key,val) + os.linesep)

print(">>>")
print(">>> Write output to "+combinedPrefix+"BudgetDefs.tex!")