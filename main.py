import sys
import os
import time
import base

print('now is', time.ctime())
settings_file_path = 'settings//settings.xml'

if os.path.exists(settings_file_path):
    settings = base.Settings(settings_file_path)
    work_dir = settings.work_dir
    source_dir = settings.source_dir
else:
    print('Файл с настройками не найден')
    sys.exit(0)

#Ищем файл подходящий под параметры ответа из ТФОМС
if not settings.search_work_files():
    sys.exit(0)

if (settings.work_dir.find('не содержит') != -1 or
    settings.source_dir.find('не содержит') != -1 or
    settings.code_lpu_for_seach_files.find('не содержит') != -1):
    print('Файл с настройками поврежден')
    sys.exit(0)

smo01 = smo10 = smo13 = smo17 = smo22 = smo24 = 0

if settings.bill_go:
    source = []
    path_to_work_dirs = os.listdir(path=str(settings.work_dir))
    path_to_work_dirs.sort()
    for name_dir in path_to_work_dirs:
        for name_file in os.listdir(path=str(settings.work_dir + name_dir)):
            if name_file.find('HM') != -1:
                file = settings.work_dir + name_dir + '/' + name_file
                source.append(file)
    
    for i in range (len(source)):
        print(source[i])
        if str(source[i]).find('S61001') != -1:
            smo01 = base.DataSMO(source[i], settings)
        elif str(source[i]).find('S61010') != -1:
            smo10 = base.DataSMO(source[i], settings)
        elif str(source[i]).find('S61013') != -1:
            smo13 = base.DataSMO(source[i], settings)
        elif str(source[i]).find('S61017') != -1:
            smo17 = base.DataSMO(source[i], settings)
        elif str(source[i]).find('S61022') != -1:
            smo22 = base.DataSMO(source[i], settings)
        elif str(source[i]).find('S61024') != -1:
            smo24 = base.DataSMO(source[i], settings)
    del(source)
    list_consolidated_ambulance_smo = []
    list_consolidated_hospital_ks_smo = []
    list_consolidated_hospital_ds_smo = []
    if smo01:
        print(smo01.bill_data['kod_smo'])
        bill = base.BillGenerator(settings=settings, bill_data=smo01.bill_data)
        if len(smo01.consolidated_ambulance_insurance_company['kod_lpu']) > 0:
            list_consolidated_ambulance_smo.append(smo01.consolidated_ambulance_insurance_company)
            consolidated_amb = base.ConsolidatedAmbulanceBillInSmoGenerator(settings=settings, consolidated_insurance_company=smo01.consolidated_ambulance_insurance_company)
        if len(smo01.consolidated_ks_insurance_company['kod_lpu']) > 0:
            list_consolidated_hospital_ks_smo.append(smo01.consolidated_ks_insurance_company)
            consolidated_hospytal = base.ConsolidatedHospitalBillInSmoGenerator(settings=settings, consolidated_insurance_company=smo01.consolidated_ks_insurance_company)
        if len(smo01.consolidated_ds_insurance_company['kod_lpu']) > 0:
            list_consolidated_hospital_ds_smo.append(smo01.consolidated_ds_insurance_company)
            consolidated_hospytal = base.ConsolidatedHospitalBillInSmoGenerator(settings=settings, consolidated_insurance_company=smo01.consolidated_ds_insurance_company)

    if smo10:
        print(smo10.bill_data['kod_smo'])
        bill = base.BillGenerator(settings=settings, bill_data=smo10.bill_data)
        if len(smo10.consolidated_ambulance_insurance_company['kod_lpu']) > 0:
            list_consolidated_ambulance_smo.append(smo10.consolidated_ambulance_insurance_company)
            consolidated_amb = base.ConsolidatedAmbulanceBillInSmoGenerator(settings=settings, consolidated_insurance_company=smo10.consolidated_ambulance_insurance_company)
        if len(smo10.consolidated_ks_insurance_company['kod_lpu']) > 0:
            list_consolidated_hospital_ks_smo.append(smo10.consolidated_ks_insurance_company)
            consolidated_hospytal = base.ConsolidatedHospitalBillInSmoGenerator(settings=settings, consolidated_insurance_company=smo10.consolidated_ks_insurance_company)
        if len(smo10.consolidated_ds_insurance_company['kod_lpu']) > 0:
            list_consolidated_hospital_ds_smo.append(smo10.consolidated_ds_insurance_company)
            consolidated_hospytal = base.ConsolidatedHospitalBillInSmoGenerator(settings=settings, consolidated_insurance_company=smo10.consolidated_ds_insurance_company)

    if smo13:
        print(smo13.bill_data['kod_smo'])
        bill = base.BillGenerator(settings=settings, bill_data=smo13.bill_data)
        if len(smo13.consolidated_ambulance_insurance_company['kod_lpu']) > 0:
            list_consolidated_ambulance_smo.append(smo13.consolidated_ambulance_insurance_company)
            consolidated_amb = base.ConsolidatedAmbulanceBillInSmoGenerator(settings=settings, consolidated_insurance_company=smo13.consolidated_ambulance_insurance_company)
        if len(smo13.consolidated_ks_insurance_company['kod_lpu']) > 0:
            list_consolidated_hospital_ks_smo.append(smo13.consolidated_ks_insurance_company)
            consolidated_hospytal = base.ConsolidatedHospitalBillInSmoGenerator(settings=settings, consolidated_insurance_company=smo13.consolidated_ks_insurance_company)
        if len(smo13.consolidated_ds_insurance_company['kod_lpu']) > 0:
            list_consolidated_hospital_ds_smo.append(smo13.consolidated_ds_insurance_company)
            consolidated_hospytal = base.ConsolidatedHospitalBillInSmoGenerator(settings=settings, consolidated_insurance_company=smo13.consolidated_ds_insurance_company)

    if smo17:
        print(smo17.bill_data['kod_smo'])
        bill = base.BillGenerator(settings=settings, bill_data=smo17.bill_data)
        if len(smo17.consolidated_ambulance_insurance_company['kod_lpu']) > 0:
            list_consolidated_ambulance_smo.append(smo17.consolidated_ambulance_insurance_company)
            consolidated_amb = base.ConsolidatedAmbulanceBillInSmoGenerator(settings=settings, consolidated_insurance_company=smo17.consolidated_ambulance_insurance_company)
        if len(smo17.consolidated_ks_insurance_company['kod_lpu']) > 0:
            list_consolidated_hospital_ks_smo.append(smo17.consolidated_ks_insurance_company)
            consolidated_hospytal = base.ConsolidatedHospitalBillInSmoGenerator(settings=settings, consolidated_insurance_company=smo17.consolidated_ks_insurance_company)
        if len(smo17.consolidated_ds_insurance_company['kod_lpu']) > 0:
            list_consolidated_hospital_ds_smo.append(smo17.consolidated_ds_insurance_company)
            consolidated_hospytal = base.ConsolidatedHospitalBillInSmoGenerator(settings=settings, consolidated_insurance_company=smo17.consolidated_ds_insurance_company)

    if smo22:
        print(smo22.bill_data['kod_smo'])
        bill = base.BillGenerator(settings=settings, bill_data=smo22.bill_data)
        if len(smo22.consolidated_ambulance_insurance_company['kod_lpu']) > 0:
            list_consolidated_ambulance_smo.append(smo22.consolidated_ambulance_insurance_company)
            consolidated_amb = base.ConsolidatedAmbulanceBillInSmoGenerator(settings=settings, consolidated_insurance_company=smo22.consolidated_ambulance_insurance_company)
        if len(smo22.consolidated_ks_insurance_company['kod_lpu']) > 0:
            list_consolidated_hospital_ks_smo.append(smo22.consolidated_ks_insurance_company)
            consolidated_hospytal = base.ConsolidatedHospitalBillInSmoGenerator(settings=settings, consolidated_insurance_company=smo22.consolidated_ks_insurance_company)
        if len(smo22.consolidated_ds_insurance_company['kod_lpu']) > 0:
            list_consolidated_hospital_ds_smo.append(smo22.consolidated_ds_insurance_company)
            consolidated_hospytal = base.ConsolidatedHospitalBillInSmoGenerator(settings=settings, consolidated_insurance_company=smo22.consolidated_ds_insurance_company)

    if smo24:
        print(smo24.bill_data['kod_smo'])
        bill = base.BillGenerator(settings=settings, bill_data=smo24.bill_data)
        if len(smo24.consolidated_ambulance_insurance_company['kod_lpu']) > 0:
            list_consolidated_ambulance_smo.append(smo24.consolidated_ambulance_insurance_company)
            consolidated_amb = base.ConsolidatedAmbulanceBillInSmoGenerator(settings=settings, consolidated_insurance_company=smo24.consolidated_ambulance_insurance_company)
        if len(smo24.consolidated_ks_insurance_company['kod_lpu']) > 0:
            list_consolidated_hospital_ks_smo.append(smo24.consolidated_ks_insurance_company)
            consolidated_hospytal = base.ConsolidatedHospitalBillInSmoGenerator(settings=settings, consolidated_insurance_company=smo24.consolidated_ks_insurance_company)
        if len(smo24.consolidated_ds_insurance_company['kod_lpu']) > 0:
            list_consolidated_hospital_ds_smo.append(smo24.consolidated_ds_insurance_company)
            consolidated_hospytal = base.ConsolidatedHospitalBillInSmoGenerator(settings=settings, consolidated_insurance_company=smo24.consolidated_ds_insurance_company)

    if len(list_consolidated_ambulance_smo) > 0:
        base.ConsolidatedAmbulanceBillGenerator(settings=settings, list_consolidated_smo=list_consolidated_ambulance_smo)
    if len(list_consolidated_hospital_ks_smo) > 0:
        base.ConsolidatedHospitalBillGenerator(settings=settings, list_consolidated_smo=list_consolidated_hospital_ks_smo)
    if len(list_consolidated_hospital_ds_smo) > 0:
        base.ConsolidatedHospitalBillGenerator(settings=settings, list_consolidated_smo=list_consolidated_hospital_ds_smo)
    
else:
    print('Нечего выполнять')

print('now is', time.ctime())
