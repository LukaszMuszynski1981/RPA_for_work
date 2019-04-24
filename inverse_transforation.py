import xml.etree.ElementTree
import re
import os
import glob
from datetime import datetime

member_pattern = r'\.[a-zA-Z]+\d*[_a-zA-Z]+\d*'
expression_pattern = r"'[a-zA-Z]+'|'[0-9]+'"
last_element_expression = r'[._a-zA-Z0-9]+[)]'
or_and_pattern = r'\.[a-z]+\.'  
in_par_pattern = r'[i-n]+\s?\([^:alnum:]*\)'
src_member = 'cpaSourceExpression'
tgt_member = 'targetMemberName'
src_object = 'sourceDataDictionaryObjectName'
tgt_object = 'targetDataDictionaryObjectName'


def add_log(text):
    
    with open(import_path + r'\{}'.format(log_file_name), 'a') as file:
        file.write('Not fully matched: {} \n'.format(text))


def inverse_condition(text, target_member, if_indicator=False, decode_indicator=False, strtran_indicator=False):
    
    in_exp = re.findall(in_par_pattern, text)
    if in_exp:
        text = text.replace(in_exp[0], "= 'shouldBeCheck'")
    
    text_split = text.split(',')
    tgt_member_copy = target_member
    final_cond_list = {}
   
    if if_indicator:
        for _ in range(0, len(text_split) - 1, 2):
            if not re.findall(or_and_pattern, text_split[_]):
                try:
                    fnd = re.findall(member_pattern, text_split[_])[0][1:]
                    target_member = fnd
                    gst_exp_fnd = re.findall(expression_pattern, text_split[_])[0]
                    last_elem = re.findall(last_element_expression, text_split[-1])[0]
                    gst_member_exp = text_split[_].replace(fnd, tgt_member_copy)
                    gst_member_exp = gst_member_exp.replace(gst_exp_fnd, text_split[_+1])
                    working_condition = ','.join([gst_member_exp, gst_exp_fnd]).strip()
                    
                    if target_member not in final_cond_list.keys():
                        final_condition = ','.join([working_condition, last_elem])
                        final_cond_list[target_member] = final_condition
                    else:
                        working_condition = ','.join([working_condition, last_elem])
                        final_condition = final_cond_list[target_member].replace('null', working_condition)
                        final_cond_list[target_member] = final_condition
    
                except IndexError:
                    add_log(text)

    if decode_indicator:
        text_split[-1] = text_split[-1].strip().replace(')', '')
        text_split[-2] = ''.join([text_split[-2], ')'])
        
        for _ in range(1, len(text_split), 2):
            text_split[_].strip()
            text_split[_ + 1].strip()
            text_split[_], text_split[_+1] = text_split[_+1], text_split[_]
        
        joined_text = ','.join([i for i in text_split])
        fnd = re.findall(member_pattern, joined_text)[0][1:]
        final_condition = joined_text.replace(fnd, tgt_member_copy)
        final_cond_list[fnd] = final_condition
     
    return list(final_cond_list.items())
    

def check_all_members(root, tag, src_member):
    
    if_indicator, decode_indicator, strtran_indicator = False, False, False

    member_dict = {}
    for child in root.iter(tag):
        try:
            member = re.findall(member_pattern, child.attrib[src_member])
            
            if '.or.' in child.attrib[src_member] or '.and.' in child.attrib[src_member]:
                add_log(child.attrib[src_member])
            
            if len(child.attrib[src_member].split('(')) > 1 and child.attrib[src_member].upper().startswith('IF'): 

                if_indicator = True

                iv = inverse_condition(child.attrib[src_member], child.attrib[tgt_member], if_indicator, decode_indicator, strtran_indicator)
                for _ in range(len(iv)):
                    member_dict[iv[_][0].strip()] = iv[_][1].strip()

                if_indicator = False
                
            if len(child.attrib[src_member].split('(')) > 1 and child.attrib[src_member].upper().startswith('DECODE'):

                decode_indicator = True

                iv = inverse_condition(child.attrib[src_member], child.attrib[tgt_member], if_indicator, decode_indicator, strtran_indicator)
                member_dict[iv[0][0].strip()] = iv[0][1].strip()

                decode_indicator = False

            if 'STRTRAN' in child.attrib[src_member]:
                
                member_dict[member[0][1:].strip()] = "if(upper({}.{})='NA', 'N/A', {}.{}))".format(root.attrib[tgt_object],child.attrib[tgt_member].strip(), root.attrib[tgt_object], child.attrib[tgt_member].strip())

            if member[0][1:] not in ['or', 'and'] and len(member) == 1 and not len(child.attrib[src_member].split('(')) > 1:
                
                if member[0][1:] in ['VOL_REV', 'VOL_ASSET', 'VOL_PREM', 'VOL_OPBUD']:
                    member_dict[member[0][1:].strip()] = child.attrib[src_member].replace('*', '/').replace('ORGDATA.', '').strip()
                else:
                    member_dict[member[0][1:].strip()] = child.attrib[tgt_member].strip()
        
        except IndexError:
            pass
        
    return list(member_dict.items())


def attributes_dict(gst_object, raw_cpa_source_expression, target_member_name, is_update_key_mapping="False"):

    if raw_cpa_source_expression.upper().startswith('IF') or raw_cpa_source_expression.upper().startswith('DECODE'):
        cpa_source_expression = raw_cpa_source_expression.replace(root.attrib[tgt_object], root.attrib[src_object])
    else:
        cpa_source_expression = '.'.join([gst_object, raw_cpa_source_expression])
    
    attrib_dict = {'version':"1", 
            'cpaSourceExpression': cpa_source_expression,
            'defaultValue':"", 
            'formattedCpaSourceExpression':"", 
            'isUpdateKeyMapping': is_update_key_mapping,
            'sourceExpressionDataType':"CpaUnknown",
            'targetMemberName': target_member_name
        }
    return attrib_dict


def inverse_object_schema_properties(root):
    
    '''zamiana obiektów source <--> target miejscami'''
    
    temp_value = root.attrib[src_object]
    root.attrib[src_object] = root.attrib[tgt_object]
    root.attrib[tgt_object] = temp_value
    root.attrib['navigatorName'] = ' '.join([root.attrib[src_object], 'to', root.attrib[tgt_object]])
    root.attrib['targetSurveyCatalogCode'] = 'CAT_GST_PR_LA'
    root.attrib['targetSurveyDictionaryCode'] = 'LAPRO_9919'
    
    for _ in root.iter('sourceDataDictionaryObjectFilter'):
        _.attrib['expression'] = ''
    
    if root.attrib[tgt_object] == 'POSDATA':
        root.attrib['transformType'] = 'Append'
    else:
        root.attrib['transformType'] = 'UpdateAndAppend'
    
    return 


run_start = datetime.now().strftime('%d_%b_%H_%M')
log_file_name = 'Log_file_{}.txt'.format(run_start)
print(run_start, log_file_name)
export_file_path = input()


try:
    os.mkdir('{}'.format(export_file_path) + r'\files_to_import')
    import_path = export_file_path + r'\files_to_import'

except FileExistsError:
    import_path = export_file_path + r'\files_to_import'


for file_path in glob.glob(os.path.join(export_file_path, '*xml')):

    tree = xml.etree.ElementTree.parse(r'{}'.format(file_path))
    root = tree.getroot() 
    
    inverse_object_schema_properties(root)
    member_list = check_all_members(root, 'mapping', src_member)

    for mapping in root.findall('mappings'):
        mapping.clear()

    for _ in range(len(member_list)):
        if member_list[_][0] in ['CPY_CODE', 'GRP_CODE', 'CTRY_CODE', 'LEVEL_CODE'] and \
                        root.attrib[tgt_object] not in ['POSDATA', 'LTIPLAN']:
            el = root.makeelement('mapping', attributes_dict(
                root.attrib[src_object],
                member_list[_][1],
                member_list[_][0],
                "True"
            ))
        else:
            el = root.makeelement('mapping', attributes_dict(
                root.attrib[src_object],
                member_list[_][1],
                member_list[_][0]
            ))
        mapping.append(el)       
    
        tree.write(import_path + r'\{}_{}_{}_to_import.xml'.format(
            file_path.split('\\')[-1].split('_')[0],
            root.attrib[src_object],
            root.attrib[tgt_object]
        ))
                          
print('check file')

