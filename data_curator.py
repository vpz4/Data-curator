# -*- coding: utf-8 -*-
"""
Created on Wed Sep  5 17:17:59 2018

@author: vpezoulas
"""

import Orange
import numpy as np
import pandas as pd
import xlwt
import re
import sys
import io
import json
import scipy
import seaborn as sns
import matplotlib.pyplot as plt
import timeit
import itertools
import webdav.client as wc
import xml.etree.ElementTree as ET
import os
from scipy.stats import spearmanr
from flask import Flask, jsonify, request
from io import StringIO
from outliers import smirnov_grubbs as grubbs
from Orange.preprocess import Impute, Average
from collections import Counter
from nltk.corpus import wordnet as wn
from Levenshtein import jaro
from difflib import SequenceMatcher
from xlrd import open_workbook
import hdbscan
import random
from flask import redirect, url_for, render_template, send_file
import zipfile
import copy
from werkzeug.utils import secure_filename
import time

#application specs
UPLOAD_FOLDER = os.path.dirname(os.path.abspath(__file__))+'/uploaded/';
DOWNLOAD_FOLDER = os.path.dirname(os.path.abspath(__file__))+'/downloaded/';
ALLOWED_EXTENSIONS = set(['xlsx','xls']);

#application specs
app = Flask(__name__);
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER;
app.config['DOWNLOAD_FOLDER'] = DOWNLOAD_FOLDER;
app.secret_key = 'some_secret'

def connect_to_webdav(user_id, pwd):
    #create options
    options = {'webdav_hostname': "http://192.168.50.6/hcloud",
               'webdav_login': user_id,
               'webdav_password': pwd};   
               
    #connect
    client = wc.Client(options);
    return client;


#capture print screen
class Capturing(list):
    def __enter__(self):
        self._stdout = sys.stdout;
        sys.stdout = self._stringio = StringIO();
        return self;
    def __exit__(self, *args):
        self.extend(self._stringio.getvalue().splitlines());
        del self._stringio;  # free up some memory
        sys.stdout = self._stdout;


#assistance for json interlinking
def mangle(s):
    return s.strip()[1:-1];


#connect json files    
def cat_json(output_filename, input_filenames):
    with open(output_filename, "w") as outfile:
        first = True
        for infile_name in input_filenames:
            with open(infile_name) as infile:
                if first:
                    outfile.write('[')
                    first = False
                else:
                    outfile.write(',')
                outfile.write(mangle(infile.read()))
        outfile.write(']')
                        
                        
def formatNumber(num):
  if num % 1 == 0:
    return int(num);
  else:
    return num;

  
def formatNumber_v2(num):
  if num % 1 == 0:
    return [int(num), 1];
  else:
    return [num, 0];


def formatNumber_v3(num):
    try:
        y = float(num);
        if(y % 1 == 0):
            return int(y);
        else:
            return y;
    except:
        return num;


def intersect(seq1, seq2):
    res = []                     # start empty
    for x in seq1:               # scan seq1
        if x in seq2:            # common item?
            res.append(x)        # add to end
    return res


def create_wr_io(path_name, pythonDictionary):
    with io.open(path_name, 'w', encoding='utf-8') as f:
        f.write(json.dumps(pythonDictionary, ensure_ascii=True, sort_keys=False, indent=4));
    
    with open(path_name) as json_data:
        d = json.load(json_data);
    
    return d;


def outliers_iqr(ys):
    [quartile_1, quartile_3] = np.percentile(ys, [25, 75]);
    iqr = quartile_3 - quartile_1;
    lower_bound = quartile_1 - (iqr * 1.5);
    upper_bound = quartile_3 + (iqr * 1.5);
    outliers_ind = np.where((ys > upper_bound) | (ys < lower_bound));
    return [iqr, outliers_ind];


def outliers_z_score(ys):
    threshold = 3;
    mean_y = np.mean(ys);
    stdev_y = np.std(ys);
    z_scores = [(y - mean_y) / stdev_y for y in ys];
    outliers_ind = np.where(np.abs(z_scores) > threshold);
    return [z_scores, outliers_ind];


def outliers_mad(ys):
    mad = np.median(np.abs(ys-np.median(ys)));
    mad_scores = [(y - mad) for y in ys];
    outliers_ind = np.where(np.abs(mad_scores) > 3*mad);
    return [mad_scores, outliers_ind];


def write_evaluation_report(data_org, r, c, features_total, metas_features, pos_metas, ranges, var_type_final, var_type_final_2, 
                            var_type_metas, var_type_metas_2, features_state_metas, incompatibilities_metas, features_missing_values_metas, 
                            bad_features_metas, bad_features_ind_metas,fair_features_metas, fair_features_ind_metas, 
                            good_features_metas,good_features_ind_metas, a_total_metas, outliers_ind_metas, y_score_metas, 
                            outliers_pos_metas, ranges_metas, features_missing_values, features_state, outliers_ind, incompatibilities, a_total, path_f):
    
    book = xlwt.Workbook(encoding="utf-8");
    sheet1 = book.add_sheet("Sheet 1");
    
    borders = xlwt.Borders();
    borders.top = 1;
    borders.bottom = 1;
    borders.left = 1;
    borders.right = 1;
    
    #header
    font0 = xlwt.Font();
    font0.name = 'Arial';
    font0.colour_index = xlwt.Style.colour_map['white'];
    font0.bold = True;
    font0.height = 280;
    
    #sub-header
    font1_1 = xlwt.Font();
    font1_1.name = 'Arial';
    font1_1.colour_index = xlwt.Style.colour_map['black'];
    font1_1.bold = True;
    font1_1.height = 220;
    
    font1_2 = font1_1;
    
    #context
    font2_1 = xlwt.Font();
    font2_1.name = 'Arial';
    font2_1.colour_index = xlwt.Style.colour_map['black'];
    font2_1.height = 220;
    
    #context with coloring
    font2_2 = font2_1;

    #context with coloring
    font2_3 = font2_1;

    #context with coloring
    font2_4 = font2_1;
    
    style0 = xlwt.XFStyle();
    style1_1 = xlwt.XFStyle();
    style1_2 = xlwt.XFStyle();
    style2_1 = xlwt.XFStyle();
    style2_1_0 = xlwt.XFStyle();
    style2_2 = xlwt.XFStyle();
    
    style0.font = font0; #header
    style1_1.font = font1_1; #sub-header
    style1_1.borders = borders;
    style1_2 = xlwt.easyxf('alignment: horizontal center');
    style1_2.font = font1_2; #sub-header
    style1_2.borders = borders;
    style2_1 = xlwt.easyxf('alignment: horizontal center');
    style2_1.borders = borders;
    style2_1.font = font2_1; #context
    style2_1_0.borders = borders;
    style2_1_0.font = font2_1; #context
    style2_2 = xlwt.easyxf('alignment: horizontal center');
    style2_2.borders = borders;
    style2_2.font = font2_2; #context with coloring
    style2_3 = xlwt.easyxf('alignment: horizontal center');
    style2_3.borders = borders;
    style2_3.font = font2_3; #context with coloring
    style2_4 = xlwt.easyxf('alignment: horizontal center');
    style2_4.borders = borders;
    style2_4.font = font2_4; #context with coloring
    
    #header
    pattern0 = xlwt.Pattern();
    pattern0.pattern = xlwt.Pattern.SOLID_PATTERN;
    pattern0.pattern_fore_colour = xlwt.Style.colour_map['blue_gray'];
    style0.pattern = pattern0;
    
    #sub-header
    pattern1 = xlwt.Pattern();
    pattern1.pattern = xlwt.Pattern.SOLID_PATTERN;
    pattern1.pattern_fore_colour = xlwt.Style.colour_map['gray25'];
    style1_1.pattern = pattern1;    
    style1_2.pattern = pattern1; 
    
    #context with coloring
    pattern2_2 = xlwt.Pattern();
    pattern2_2.pattern = xlwt.Pattern.SOLID_PATTERN;
    pattern2_2.pattern_fore_colour = xlwt.Style.colour_map['rose'];
    style2_2.pattern = pattern2_2;

    #context with coloring
    pattern2_3 = xlwt.Pattern();
    pattern2_3.pattern = xlwt.Pattern.SOLID_PATTERN;
    pattern2_3.pattern_fore_colour = xlwt.Style.colour_map['light_green'];
    style2_3.pattern = pattern2_3;

    #context with coloring
    pattern2_4 = xlwt.Pattern();
    pattern2_4.pattern = xlwt.Pattern.SOLID_PATTERN;
    pattern2_4.pattern_fore_colour = xlwt.Style.colour_map['periwinkle'];
    style2_4.pattern = pattern2_4;
    
    offs_1 = 0;
    offs_2 = 0;
    
    #panel_1
    
    n_disc = Counter(var_type_final_2+var_type_metas_2)['categorical'];
    n_cont = Counter(var_type_final_2+var_type_metas_2)['numeric'];
    n_unkn = Counter(var_type_final_2+var_type_metas_2)['unknown'];
    n_miss = np.around(100*((a_total+a_total_metas)/((c+len(metas_features))*r)), 2);
    
    sheet1.write_merge(offs_1,offs_1,offs_2,offs_2+4,"Metadata",style0);
    sheet1.write_merge(offs_1+1,offs_1+1,offs_2,offs_2+3,"Number of feature(s)",style1_1); sheet1.write(offs_1+1,offs_2+4,np.str(c+len(metas_features)),style2_1);
    sheet1.write_merge(offs_1+2,offs_1+2,offs_2,offs_2+3,"Number of instance(s)",style1_1); sheet1.write(offs_1+2,offs_2+4,np.str(r),style2_1);
    sheet1.write_merge(offs_1+3,offs_1+3,offs_2,offs_2+3,"Discrete feature(s)",style1_1); sheet1.write(offs_1+3,offs_2+4,np.str(n_disc),style2_1);
    sheet1.write_merge(offs_1+4,offs_1+4,offs_2,offs_2+3,"Continuous feature(s)",style1_1); sheet1.write(offs_1+4,offs_2+4,np.str(n_cont),style2_1);
    sheet1.write_merge(offs_1+5,offs_1+5,offs_2,offs_2+3,"Unknown feature(s)",style1_1); sheet1.write(offs_1+5,offs_2+4,np.str(n_unkn),style2_1);
    sheet1.write_merge(offs_1+6,offs_1+6,offs_2,offs_2+3,"Missing values (%)",style1_1); sheet1.write(offs_1+6,offs_2+4,np.str(n_miss)+"%",style2_1);
    
    #panel_2
    sheet1.write_merge(offs_1+8,offs_1+8,offs_2,offs_2+22,"Quality assessment",style0);
    c1=0;
    sheet1.write_merge(c1+9,c1+9,offs_2,offs_2+4,"Features",style1_2);
    for i in range(c+len(metas_features)):
        sheet1.write_merge(c1+10,c1+10,offs_2,offs_2+4,features_total[i],style2_1_0);
        c1=c1+1;
        
    c1=0; c2=0; c3=0;
    sheet1.write_merge(c1+9,c1+9,offs_2+5,offs_2+8,"Value range",style1_2);
    for i in range(c+len(metas_features)):
        if(i not in pos_metas):
            ranges_f = str(ranges[c2]).replace("'",'');
            sheet1.write_merge(c1+10,c1+10,offs_2+5,offs_2+8,ranges_f,style2_1); c2 = c2+1;
        else:
#            ranges_f = [x for x in ranges_metas[c3] if x];
            ranges_f = str(ranges_metas[c3]).replace("'",'');
            sheet1.write_merge(c1+10,c1+10,offs_2+5,offs_2+8,ranges_f,style2_1); c3 = c3+1;
        c1=c1+1;
        
    c1=0; c2=0; c3=0;
    sheet1.write_merge(c1+9,c1+9,offs_2+9,offs_2+10,"Type",style1_2);
    for i in range(c+len(metas_features)):
        if(i not in pos_metas):
            sheet1.write_merge(c1+10,c1+10,offs_2+9,offs_2+10,var_type_final_2[c2],style2_1); c2 = c2+1;
        else:
            sheet1.write_merge(c1+10,c1+10,offs_2+9,offs_2+10,var_type_metas_2[c3],style2_1);
        c1=c1+1;
        
    c1=0; c2=0; c3=0;
    sheet1.write_merge(c1+9,c1+9,offs_2+11,offs_2+12,"Variable type",style1_2); 
    for i in range(c+len(metas_features)):
        if(i not in pos_metas):
            sheet1.write_merge(c1+10,c1+10,offs_2+11,offs_2+12,var_type_final[c2],style2_1); c2 = c2+1;
        else:
            sheet1.write_merge(c1+10,c1+10,offs_2+11,offs_2+12,var_type_metas[c3],style2_1); c3 = c3+1;
        c1=c1+1;
        
    c1=0; c2=0; c3=0;
    sheet1.write_merge(c1+9,c1+9,offs_2+13,offs_2+15,"Missing values",style1_2);
    for i in range(c+len(metas_features)):
        if(i not in pos_metas):
            sheet1.write_merge(c1+10,c1+10,offs_2+13,offs_2+15,features_missing_values[c2],style2_1); c2 = c2+1;
        else:
            sheet1.write_merge(c1+10,c1+10,offs_2+13,offs_2+15,features_missing_values_metas[c3],style2_1); c3 = c3+1;
        c1=c1+1;

    c1=0; c2=0; c3=0;
    sheet1.write_merge(c1+9,c1+9,offs_2+16,offs_2+17,"State",style1_2);
    for i in range(c+len(metas_features)):
        if(i not in pos_metas):
            if(features_state[c2] == "bad"):
                sheet1.write_merge(c1+10,c1+10,offs_2+16,offs_2+17,features_state[c2],style2_2);
            elif(features_state[c2] == "fair"):
                sheet1.write_merge(c1+10,c1+10,offs_2+16,offs_2+17,features_state[c2],style2_3);
            elif(features_state[c2] == "good"):
                sheet1.write_merge(c1+10,c1+10,offs_2+16,offs_2+17,features_state[c2],style2_4);
            c2=c2+1;
        else:
            if(features_state_metas[c3] == "bad"):
                sheet1.write_merge(c1+10,c1+10,offs_2+16,offs_2+17,features_state_metas[c3],style2_2);
            elif(features_state_metas[c3] == "fair"):
                sheet1.write_merge(c1+10,c1+10,offs_2+16,offs_2+17,features_state_metas[c3],style2_3);
            elif(features_state_metas[c3] == "good"):
                sheet1.write_merge(c1+10,c1+10,offs_2+16,offs_2+17,features_state_metas[c3],style2_4);               
            c3=c3+1;
        c1=c1+1;
        
    c1=0; c2=0; c3=0;
    sheet1.write_merge(c1+9,c1+9,offs_2+18,offs_2+19,"Outliers",style1_2);
    for i in range(c+len(metas_features)):
        if(i not in pos_metas):
            if(outliers_ind[c2] == "yes"):
                sheet1.write_merge(c1+10,c1+10,offs_2+18,offs_2+19,outliers_ind[c2],style2_2);
            elif(outliers_ind[c2] == "no"):
                sheet1.write_merge(c1+10,c1+10,offs_2+18,offs_2+19,outliers_ind[c2],style2_4);
            elif(outliers_ind[c2] == "not-applicable"):
                sheet1.write_merge(c1+10,c1+10,offs_2+18,offs_2+19,outliers_ind[c2],style2_2);
            c2=c2+1;    
        else:
            if(outliers_ind_metas[c3] == "yes"):
                sheet1.write_merge(c1+10,c1+10,offs_2+18,offs_2+19,outliers_ind_metas[c3],style2_2);
            elif(outliers_ind_metas[c3] == "no"):
                sheet1.write_merge(c1+10,c1+10,offs_2+18,offs_2+19,outliers_ind_metas[c3],style2_4);
            elif(outliers_ind_metas[c3] == "not-applicable"):
                sheet1.write_merge(c1+10,c1+10,offs_2+18,offs_2+19,outliers_ind_metas[c3],style2_2);
            c3=c3+1;                
        c1=c1+1;

    c1=0; c2=0; c3=0;
    sheet1.write_merge(c1+9,c1+9,offs_2+20,offs_2+22,"Incompatibilities",style1_2);
    for i in range(c+len(metas_features)):
        if(i not in pos_metas):
            if(incompatibilities[c2] == "no"):
                if(features_state[c2] == "bad"):
                    sheet1.write_merge(c1+10,c1+10,offs_2+20,offs_2+22,'yes, bad feature',style2_2);
                else:
                    sheet1.write_merge(c1+10,c1+10,offs_2+20,offs_2+22,incompatibilities[c2],style2_4);
            elif(incompatibilities[c2] == "yes, unknown type of data"):
                sheet1.write_merge(c1+10,c1+10,offs_2+20,offs_2+22,incompatibilities[c2],style2_2);
            c2=c2+1;
        else:
            if(incompatibilities_metas[c3] == "no"):
                if(features_state_metas[c3] == "bad"):
                    sheet1.write_merge(c1+10,c1+10,offs_2+20,offs_2+22,'yes, bad feature',style2_2);
                else:
                    sheet1.write_merge(c1+10,c1+10,offs_2+20,offs_2+22,incompatibilities_metas[c3],style2_4);
            elif(incompatibilities_metas[c3] == "yes, unknown type of data"):
                sheet1.write_merge(c1+10,c1+10,offs_2+20,offs_2+22,incompatibilities_metas[c3],style2_2);
            c3=c3+1;            
        c1=c1+1;
    
    if not os.path.exists('results'):
        os.makedirs('results');
        
    book.save(path_f);


def write_curated_dataset(data_org, wb, pos_metas, features_total, features_state, metas_features, imputation_method_id, 
                          var_type_final, outliers_pos, r, c, features_state_metas, outliers_pos_metas, var_type_metas, 
                          var_type_metas_2, y_total, incomp_pos, incomp_pos_metas, path_f):
    
    book = xlwt.Workbook(encoding="utf-8");
    sheet1 = book.add_sheet("Sheet 1");

    borders = xlwt.Borders();
    borders.top = 1;
    borders.bottom = 1;
    borders.left = 1;
    borders.right = 1;
    
    #headers
    font1 = xlwt.Font();
    font1.name = 'Arial';
    font1.colour_index = xlwt.Style.colour_map['black'];
    font1.height = 200;
    
    style1 = xlwt.XFStyle(); #outliers
    style2 = xlwt.XFStyle(); #imputed values
    style3 = xlwt.XFStyle(); #bad feature
    style4 = xlwt.XFStyle(); #fair feature
    style5 = xlwt.XFStyle(); #good feature
    style6 = xlwt.XFStyle(); #incomp values
    
    pattern1 = xlwt.Pattern(); #outliers
    pattern2 = xlwt.Pattern(); #imputed values
    pattern3 = xlwt.Pattern(); #bad feature
    pattern4 = xlwt.Pattern(); #fair feature
    pattern5 = xlwt.Pattern(); #good feature
    pattern6 = xlwt.Pattern(); #incomp values
    
    #outliers
    pattern1.pattern = xlwt.Pattern.SOLID_PATTERN;
    pattern1.pattern_fore_colour = xlwt.Style.colour_map['gold'];
    style1.pattern = pattern1;
    style1.font = font1;
    
    #imputed values
    pattern2.pattern = xlwt.Pattern.SOLID_PATTERN;
    pattern2.pattern_fore_colour = xlwt.Style.colour_map['gray25'];
    style2.pattern = pattern2;
    style2.font = font1;

    #bad feature
    pattern3.pattern = xlwt.Pattern.SOLID_PATTERN;
    pattern3.pattern_fore_colour = xlwt.Style.colour_map['rose'];
    style3.pattern = pattern3;
    style3.font = font1;
    
    #fair feature
    pattern4.pattern = xlwt.Pattern.SOLID_PATTERN;
    pattern4.pattern_fore_colour = xlwt.Style.colour_map['light_green'];
    style4.pattern = pattern4;
    style4.font = font1;

    #good feature
    pattern5.pattern = xlwt.Pattern.SOLID_PATTERN;
    pattern5.pattern_fore_colour = xlwt.Style.colour_map['ice_blue'];
    style5.pattern = pattern5;
    style5.font = font1;

    #incomp values
    pattern6.pattern = xlwt.Pattern.SOLID_PATTERN;
    pattern6.pattern_fore_colour = xlwt.Style.colour_map['coral'];
    style6.pattern = pattern6;
    style6.font = font1;
    
    if(imputation_method_id == 1):
        imputer = Impute(method=Average());
        my_data = imputer(data_org);
        for j in range(c):
            if(var_type_final[j] == 'date')|(var_type_final[j] == 'int'):
                v = np.around(my_data[:,j][:]);
                if(np.max(v)>np.max(my_data[:,j][:])):
                    v = np.floor(my_data[:,j][:]);
                my_data[:,j] = v;
    elif(imputation_method_id == 2):
        T = np.isnan(data_org);
        imputer = Impute(method=Average());
        my_data = imputer(data_org);
        for j in range(c):
            f = np.where(T[:,j] == False);
            f_nan = np.where(T[:,j] == True);
            X_f = copy.deepcopy(data_org.X[:,j]);
            b = X_f[f];
            if(np.size(np.where(T[:,j] == 1)) <= np.size(b)):
                f_ind = random.sample(range(0, np.size(b)), np.size(np.where(T[:,j] == 1)));
                X_f[f_nan] = b[f_ind];
                my_data[:,j] = X_f.reshape(-1,1);
            else:
                my_data[:,j] = X_f.reshape(-1,1);
    else:
        my_data = data_org;
                        
    c1 = 0;
    c2 = 0;
    sheet_names = wb.sheet_names();
    sheet = wb.sheet_by_name(sheet_names[0]);
    for j in range(c+len(pos_metas)):
        if(j not in pos_metas):
            if(features_state[c1] == 'bad'):
                sheet1.write(0,j,features_total[j],style3);
            elif(features_state[c1] == 'fair'):
                sheet1.write(0,j,features_total[j],style4);
            elif(features_state[c1] == 'good'):
                sheet1.write(0,j,features_total[j],style5);
            sheet1.col(j).width = 256*20;  
            for i in range(r):
                if(outliers_pos[c1] != '-'):
                    if((i in outliers_pos[c1])&(str(data_org[:,c1][i]).strip()[1:-1]=='?')&(features_state[c1]=='bad')):
                        sheet1.write(i+1,j,formatNumber_v3(str(data_org[:,c1][i]).strip()[1:-1]),style1);
                    elif((i in outliers_pos[c1])&(str(data_org[:,c1][i]).strip()[1:-1]=='?')&(features_state[c1]!='bad')):
#                        sheet1.write(i+1,j,formatNumber_v3(str(my_data[:,c1][i]).strip()[1:-1]),style1);
                        sheet1.write(i+1,j,formatNumber_v3(str(data_org[:,c1][i]).strip()[1:-1]),style1);
                    elif((i in outliers_pos[c1])&(str(data_org[:,c1][i]).strip()[1:-1]!='?')&(features_state[c1]=='bad')):
                        sheet1.write(i+1,j,formatNumber_v3(str(data_org[:,c1][i]).strip()[1:-1]),style1);
                    elif((i in outliers_pos[c1])&(str(data_org[:,c1][i]).strip()[1:-1]!='?')&(features_state[c1]!='bad')):
                        sheet1.write(i+1,j,formatNumber_v3(str(data_org[:,c1][i]).strip()[1:-1]),style1);
                    elif((str(data_org[:,c1][i]).strip()[1:-1]=='?')&(features_state[c1]=='bad')&(var_type_final[c1]!='string')&(var_type_final[c1]!='unknown')):
                        sheet1.write(i+1,j,formatNumber_v3(str(data_org[:,c1][i]).strip()[1:-1]),style2);
                    elif((str(data_org[:,c1][i]).strip()[1:-1]=='?')&(features_state[c1]!='bad')):
#                        sheet1.write(i+1,j,formatNumber_v3(str(my_data[:,c1][i]).strip()[1:-1]),style2);
                        sheet1.write(i+1,j,formatNumber_v3(str(data_org[:,c1][i]).strip()[1:-1]),style2);
                    elif((str(data_org[:,c1][i]).strip()[1:-1]!='?')&(features_state[c1]=='bad')):
                        sheet1.write(i+1,j,formatNumber_v3(str(data_org[:,c1][i]).strip()[1:-1]));
                    elif((str(data_org[:,c1][i]).strip()[1:-1]!='?')&(features_state[c1]!='bad')):
                        sheet1.write(i+1,j,formatNumber_v3(str(data_org[:,c1][i]).strip()[1:-1]));
                else:
                    if((str(data_org[:,c1][i]).strip()[1:-1]=='?')&(features_state[c1]=='bad')&(var_type_final[c1]!='string')&(var_type_final[c1]!='unknown')):
                        sheet1.write(i+1,j,formatNumber_v3(str(data_org[:,c1][i]).strip()[1:-1]),style2);
                    elif((str(data_org[:,c1][i]).strip()[1:-1]=='?')&(features_state[c1]=='bad')):
                        sheet1.write(i+1,j,formatNumber_v3(str(data_org[:,c1][i]).strip()[1:-1]),style2);
                    elif((str(data_org[:,c1][i]).strip()[1:-1]=='?')&(features_state[c1]!='bad')):
                        if(incomp_pos[c1] == '-'):
                            sheet1.write(i+1,j,formatNumber_v3(str(my_data[:,c1][i]).strip()[1:-1]),style2);
                        else:
                            sheet1.write(i+1,j,formatNumber_v3(str(data_org[:,c1][i]).strip()[1:-1]),style2);
                    elif((str(data_org[:,c1][i]).strip()[1:-1]!='?')&(features_state[c1]=='bad')):
                        if(i in list(incomp_pos[c1])):
                            sheet1.write(i+1,j,formatNumber_v3(str(data_org[:,c1][i]).strip()[1:-1]),style6);
                        else:
                            sheet1.write(i+1,j,formatNumber_v3(str(data_org[:,c1][i]).strip()[1:-1]));
                    elif((str(data_org[:,c1][i]).strip()[1:-1]!='?')&(features_state[c1]!='bad')):
                        if(i in list(incomp_pos[c1])):
                            sheet1.write(i+1,j,formatNumber_v3(str(data_org[:,c1][i]).strip()[1:-1]),style6);
                        else:
                            sheet1.write(i+1,j,formatNumber_v3(str(data_org[:,c1][i]).strip()[1:-1]));                  
            c1 = c1+1;
        elif(j in pos_metas):
            if(features_state_metas[c2] == 'bad'):
                sheet1.write(0,j,features_total[j],style3);
            elif(features_state_metas[c2] == 'fair'):
                sheet1.write(0,j,features_total[j],style4);
            elif(features_state_metas[c2] == 'good'):
                sheet1.write(0,j,features_total[j],style5);
            sheet1.col(j).width = 256*20;
            for i in range(r):
                if(outliers_pos_metas[c2] != '-'):
                    if((i in outliers_pos_metas[c2])&(str(sheet.cell(i+1,j).value)=='')&(features_state_metas[c2]=='bad')):
                        sheet1.write(i+1,j,formatNumber_v3(str('?').strip()),style1);                           
                    elif((i in outliers_pos_metas[c2])&(str(sheet.cell(i+1,j).value)=='')&(features_state_metas[c2]!='bad')):
                        sheet1.write(i+1,j,formatNumber_v3(str('?').strip()),style1);                              
                    elif((i in outliers_pos_metas[c2])&(str(sheet.cell(i+1,j).value)!='')&(features_state_metas[c2]=='bad')):
                        sheet1.write(i+1,j,formatNumber_v3(str(sheet.cell(i+1,j).value).strip()),style1);
                    elif((i in outliers_pos_metas[c2])&(str(sheet.cell(i+1,j).value)!='')&(features_state_metas[c2]!='bad')):
                        sheet1.write(i+1,j,formatNumber_v3(str(sheet.cell(i+1,j).value).strip()),style1);    
                    elif((str(sheet.cell(i+1,j).value).strip()=='')&(features_state_metas[c2]=='bad')):
                        sheet1.write(i+1,j,formatNumber_v3(str('?').strip()),style2);                            
                    elif((str(sheet.cell(i+1,j).value).strip()=='')&(features_state_metas[c2]!='bad')):
                        sheet1.write(i+1,j,formatNumber_v3(str('?').strip()),style2);                              
                    elif((str(sheet.cell(i+1,j).value).strip()!='')&(features_state_metas[c2]=='bad')):
                        sheet1.write(i+1,j,formatNumber_v3(str(sheet.cell(i+1,j).value).strip()));
                    elif((str(sheet.cell(i+1,j).value).strip()!='')&(features_state_metas[c2]!='bad')):
                        sheet1.write(i+1,j,formatNumber_v3(str(sheet.cell(i+1,j).value).strip()));
                else:
                    if((str(sheet.cell(i+1,j).value).strip()=='')&(features_state_metas[c2]=='bad')&(var_type_metas[c2]!='string')&(var_type_metas[c2]!='unknown')):
                        sheet1.write(i+1,j,formatNumber_v3(str('?').strip()),style2);
                    elif((str(sheet.cell(i+1,j).value).strip()=='')&(features_state_metas[c2]=='bad')):
                        sheet1.write(i+1,j,formatNumber_v3(str('?').strip()),style2);    
                    elif((str(sheet.cell(i+1,j).value).strip()=='')&(features_state_metas[c2]!='bad')):
                        sheet1.write(i+1,j,formatNumber_v3(str('?').strip()),style2);          
                    elif((str(sheet.cell(i+1,j).value).strip()!='')&(features_state_metas[c2]=='bad')):
                        if(i in list(incomp_pos_metas[c2])):
                            sheet1.write(i+1,j,formatNumber_v3(str(sheet.cell(i+1,j).value).strip()),style6);
                        else:
                            sheet1.write(i+1,j,formatNumber_v3(str(sheet.cell(i+1,j).value).strip()));
                    elif((str(sheet.cell(i+1,j).value).strip()!='')&(features_state_metas[c2]!='bad')):
                        if(i in list(incomp_pos_metas[c2])):
                            sheet1.write(i+1,j,formatNumber_v3(str(sheet.cell(i+1,j).value).strip()),style6);
                        else:
                            sheet1.write(i+1,j,formatNumber_v3(str(sheet.cell(i+1,j).value).strip()));
            c2 = c2+1;
    
    if not os.path.exists('results'):
        os.makedirs('results');
        
    book.save(path_f);


def write_standardization_report(total_features, list_raw, list_raw_class, list_ref, list_score_class, list_match_type, list_range_ref, path_f):
    book = xlwt.Workbook(encoding="utf-8");
    sheet1 = book.add_sheet("Sheet 1");
    
    borders = xlwt.Borders();
    borders.top = 1;
    borders.bottom = 1;
    borders.left = 1;
    borders.right = 1;
    
    #header
    font0 = xlwt.Font();
    font0.name = 'Arial';
    font0.colour_index = xlwt.Style.colour_map['white'];
    font0.bold = True;
    font0.height = 280;
    
    #sub-header
    font1_1 = xlwt.Font();
    font1_1.name = 'Arial';
    font1_1.colour_index = xlwt.Style.colour_map['black'];
    font1_1.bold = True;
    font1_1.height = 220;
    
    font1_2 = font1_1;
    
    #context
    font2_1 = xlwt.Font();
    font2_1.name = 'Arial';
    font2_1.colour_index = xlwt.Style.colour_map['black'];
    font2_1.height = 220;
    
    #context with coloring
    font2_2 = font2_1;

    #context with coloring
    font2_3 = font2_1;

    #context with coloring
    font2_4 = font2_1;
    
    style0 = xlwt.XFStyle();
    style1_1 = xlwt.XFStyle();
    style1_2 = xlwt.XFStyle();
    style2_1 = xlwt.XFStyle();
    style2_1_0 = xlwt.XFStyle();
    style2_2 = xlwt.XFStyle();
    
    style0.font = font0; #header
    style1_1.font = font1_1; #sub-header
    style1_1.borders = borders;
    style1_2 = xlwt.easyxf('alignment: horizontal center');
    style1_2.font = font1_2; #sub-header
    style1_2.borders = borders;
    style2_1 = xlwt.easyxf('alignment: horizontal center');
    style2_1.borders = borders;
    style2_1.font = font2_1; #context
    style2_1_0.borders = borders;
    style2_1_0.font = font2_1; #context
    style2_2 = xlwt.easyxf('alignment: horizontal center');
    style2_2.borders = borders;
    style2_2.font = font2_2; #context with coloring
    style2_3 = xlwt.easyxf('alignment: horizontal center');
    style2_3.borders = borders;
    style2_3.font = font2_3; #context with coloring
    style2_4 = xlwt.easyxf('alignment: horizontal center');
    style2_4.borders = borders;
    style2_4.font = font2_4; #context with coloring
    
    #header
    pattern0 = xlwt.Pattern();
    pattern0.pattern = xlwt.Pattern.SOLID_PATTERN;
    pattern0.pattern_fore_colour = xlwt.Style.colour_map['blue_gray'];
    style0.pattern = pattern0;
    
    #sub-header
    pattern1 = xlwt.Pattern();
    pattern1.pattern = xlwt.Pattern.SOLID_PATTERN;
    pattern1.pattern_fore_colour = xlwt.Style.colour_map['gray25'];
    style1_1.pattern = pattern1;    
    style1_2.pattern = pattern1; 
    
    #context with coloring
    pattern2_2 = xlwt.Pattern();
    pattern2_2.pattern = xlwt.Pattern.SOLID_PATTERN;
    pattern2_2.pattern_fore_colour = xlwt.Style.colour_map['rose'];
    style2_2.pattern = pattern2_2;

    #context with coloring
    pattern2_3 = xlwt.Pattern();
    pattern2_3.pattern = xlwt.Pattern.SOLID_PATTERN;
    pattern2_3.pattern_fore_colour = xlwt.Style.colour_map['light_green'];
    style2_3.pattern = pattern2_3;

    #context with coloring
    pattern2_4 = xlwt.Pattern();
    pattern2_4.pattern = xlwt.Pattern.SOLID_PATTERN;
    pattern2_4.pattern_fore_colour = xlwt.Style.colour_map['periwinkle'];
    style2_4.pattern = pattern2_4;
    
    offs_1 = 0;
    offs_2 = 0;
    c = len(list_raw);
    
    sheet1.write_merge(offs_1,offs_1,offs_2,offs_2+21,"Data standardization report",style0);
    c1=0;
    sheet1.write_merge(c1+1,c1+1,offs_2,offs_2+4,"Features",style1_2);
    for i in range(c):
        sheet1.write_merge(c1+2,c1+2,offs_2,offs_2+4,list_raw[i],style2_1_0);
        c1=c1+1;
    
    c1=0;
    sheet1.write_merge(c1+1,c1+1,offs_2+5,offs_2+9,"Matched term or category from the reference model",style1_2);
    for i in range(c):
        sheet1.write_merge(c1+2,c1+2,offs_2+5,offs_2+9,list_ref[i],style2_1);
        c1=c1+1;

    c1=0;
    sheet1.write_merge(c1+1,c1+1,offs_2+10,offs_2+12,"Matching score",style1_2);
    for i in range(c):
        sheet1.write_merge(c1+2,c1+2,offs_2+10,offs_2+12,list_score_class[i],style2_1);
        c1=c1+1;        

    c1=0;
    sheet1.write_merge(c1+1,c1+1,offs_2+13,offs_2+15,"Type of match",style1_2);
    for i in range(c):
        sheet1.write_merge(c1+2,c1+2,offs_2+13,offs_2+15,list_match_type[i],style2_1);
        c1=c1+1;

    c1=0;
    sheet1.write_merge(c1+1,c1+1,offs_2+16,offs_2+18,"Final range",style1_2);
    for i in range(c):
        sheet1.write_merge(c1+2,c1+2,offs_2+16,offs_2+18,list_range_ref[i],style2_1);
        c1=c1+1;
        
    c1=0;
    sheet1.write_merge(c1+1,c1+1,offs_2+19,offs_2+21,"Class",style1_2);
    for i in range(c):
        sheet1.write_merge(c1+2,c1+2,offs_2+19,offs_2+21,list_raw_class[i],style2_1);
        c1=c1+1;
        
    #write book
#    path_to_save = path_f.split('/')[0]+'/'+path_f.split('/')[1]+'/'+path_f.split('/')[2];
    if not os.path.exists('results'):
        os.makedirs('results');
        
    book.save(path_f);
    

def QualityAssessment_S1(data_org, c):
    features_missing_values = [];
    bad_features = [];
    bad_features_ind = [];
    fair_features = [];
    fair_features_ind = [];
    good_features = [];
    good_features_ind = [];
    features_state = [];
    a_total = 0;
    a = np.zeros(c);
        
    for i in range(c):
        a[i] = sum(1 for d in data_org if np.isnan(d[i]));
#        a[i] = np.around((100*n_miss)/len(data_org),2);
        features_missing_values.append(a[i]);
        if(a[i]>=(len(data_org)/2)):
            bad_features.append(data_org.domain.attributes[i].name.replace('\n',' ').replace('  ', ' '));
            features_state.append('bad'); bad_features_ind.append(i);
        elif((a[i]>0) & (a[i]<(len(data_org)/2))):
            fair_features.append(data_org.domain.attributes[i].name.replace('\n',' ').replace('  ', ' '));
            features_state.append('fair'); fair_features_ind.append(i);
        elif(a[i]==0):
            good_features.append(data_org.domain.attributes[i].name.replace('\n',' ').replace('  ', ' '));
            features_state.append('good'); good_features_ind.append(i);
        a_total = a_total+a[i];
    
    return [features_missing_values, bad_features, bad_features_ind, fair_features, fair_features_ind, 
            good_features, good_features_ind, features_state, a, a_total];


def outliers_detection(data_org, c, var_type_final, outlier_detection_method_id, bad_features_ind, fair_features_ind, good_features_ind, features_missing_values):
    T = np.isnan(data_org);
    outliers_ind = [];
    y_score = [];
    outliers_pos = [];
    
    if(outlier_detection_method_id == 1):
        for j in range(c):
            if((var_type_final[j] == 'string')|(var_type_final[j] == 'unknown')):
                outliers_ind.append('not-applicable');
                y_score.append('-'); outliers_pos.append('-');
            elif(((var_type_final[j] == 'float')|(var_type_final[j] == 'date')|(var_type_final[j] == 'int'))&(features_missing_values[j] != len(data_org))):
#                if(j not in bad_features_ind):
                    f = np.where(T[:,j]==False);
                    X_f = data_org.X[:,j];
                    b = X_f[f];
                    [z_scores, outliers_ind_z_scores] = outliers_z_score(b);
                    y_score.append(np.mean(z_scores));
                    if(len(outliers_ind_z_scores[0]) != 0):
                        if(var_type_final[j] == 'int')&(np.min(b) == 0)&(np.max(b) == 1):
                            outliers_ind.append('no');
                            outliers_pos.append('-');
                        else:
                            t = list(outliers_ind_z_scores[0]);
                            outliers_pos.append(list(f[0][t]));
                            outliers_ind.append('yes');
                    else:
                        outliers_ind.append('no');
                        outliers_pos.append('-');
            elif(((var_type_final[j] == 'float')|(var_type_final[j] == 'date')|(var_type_final[j] == 'int'))&(features_missing_values[j] == len(data_org))):
                outliers_ind.append('not-applicable');
                outliers_pos.append('-');
    elif(outlier_detection_method_id == 2):
        for j in range(c):
            if((var_type_final[j] == 'string')|(var_type_final[j] == 'unknown')):
                outliers_ind.append('not-applicable');
                y_score.append('-'); outliers_pos.append('-');
            elif(((var_type_final[j] == 'float')|(var_type_final[j] == 'date')|(var_type_final[j] == 'int'))&(features_missing_values[j] != len(data_org))):
#                if(j not in bad_features_ind):
                    f = np.where(T[:,j]==False);
                    X_f = data_org.X[:,j];
                    b = X_f[f];
                    [iqr, outliers_ind_iqr] = outliers_iqr(b);
                    y_score.append(np.mean(iqr));
                    if(len(outliers_ind_iqr[0]) != 0):
                        if(var_type_final[j] == 'int')&(np.min(b) == 0)&(np.max(b) == 1):
                            outliers_ind.append('no');
                            outliers_pos.append('-');
                        else:
                            t = list(outliers_ind_iqr[0]);
                            outliers_pos.append(list(f[0][t]));
                            outliers_ind.append('yes');
                    else:
                        outliers_ind.append('no');
                        outliers_pos.append('-');
            elif(((var_type_final[j] == 'float')|(var_type_final[j] == 'date')|(var_type_final[j] == 'int'))&(features_missing_values[j] == len(data_org))):
                outliers_ind.append('not-applicable');
                outliers_pos.append('-');
    elif(outlier_detection_method_id == 3):
        for j in range(c):
            if((var_type_final[j] == 'string')|(var_type_final[j] == 'unknown')):
                outliers_ind.append('not-applicable');
                y_score.append('-'); outliers_pos.append('-');
            elif(((var_type_final[j] == 'float')|(var_type_final[j] == 'date')|(var_type_final[j] == 'int'))&(features_missing_values[j] != len(data_org))):
#                if(j not in bad_features_ind):
                    f = np.where(T[:,j]==False);
                    X_f = data_org.X[:,j];
                    b = X_f[f];
                    h_min = grubbs.min_test_indices(b);
                    h_max = grubbs.max_test_indices(b);
                    h_tol = np.unique(np.union1d(h_min, h_max));
                    y_score.append(np.mean(grubbs.test(b, alpha=0.05)));
                    if(len(h_tol) != 0):
                        if(var_type_final[j] == 'int')&(np.min(b) == 0)&(np.max(b) == 1):
                            outliers_ind.append('no');
                            outliers_pos.append('-');
                        else:
                            t = list(h_tol);
                            outliers_pos.append(list(np.take(f,t)));
                            outliers_ind.append('yes');
                    else:
                        outliers_ind.append('no');
                        outliers_pos.append('-');
    elif(outlier_detection_method_id == 4):
#        list_comp = [];
#        for j in range(c):
#            if((var_type_final[j] != 'string')&(var_type_final[j] != 'unknown')):
#                list_comp.append(j);
        
        X_c = data_org.X[:,good_features_ind]; #IMPORTANT
        
        clusterer = hdbscan.HDBSCAN(min_cluster_size=2).fit(X_c);
        y_pred_lof = clusterer.outlier_scores_; #lof_ind = np.where(y_score > 0.85)[0];
        y_score.append(np.mean(y_pred_lof));
        
        c1 = 0;
        for j in range(c):
            if(j in good_features_ind):            
                if(y_pred_lof[c1] >= 0.8):
                    outliers_ind.append('yes');
                    outliers_pos.append('-');
                else:
                    outliers_ind.append('no');
                    outliers_pos.append('-');
                c1 = c1+1;
            else:
                outliers_ind.append('not-applicable');
                outliers_pos.append('-');
    else:
        for j in range(c):
            outliers_pos.append('-');
            outliers_ind.append('no');
            y_score = [];
                
    return [outliers_ind, y_score, outliers_pos];


def metas_handler(data_org, wb, outlier_detection_method_id):
    metas = data_org.domain.metas;
    metas_features = [];
    for i in range(len(metas)):
        metas_features.append(metas[i].name);

    features_total = [];
    sheet_names = wb.sheet_names();
    sheet = wb.sheet_by_name(sheet_names[0]);
    number_of_columns = sheet.ncols;
    for col in range(number_of_columns):
        features_total.append((sheet.cell(0,col).value));            

    pos_metas = [];
    for i in range(len(features_total)):
        for j in range(len(metas_features)):
            if(jaro(metas_features[j], features_total[i]) == 1):
                pos_metas.append(i);
                
    number_of_rows = sheet.nrows;
    y_total = [];
    var_type_metas = [];
    var_type_metas_2 = [];
    incompatibilities_metas = [];
    features_state_metas = [];
    ranges_metas = [];
    incomp_pos_metas = [];
        
    #data annotation
    for j in range(len(pos_metas)):
        y_r = [];
        for i in range(1,number_of_rows):
            y = sheet.cell(i,pos_metas[j]).value;
            try:
                y_mod = float(y);
                y_mod = formatNumber(y_mod); 
                y_r.append(y_mod);
            except:
                y_r.append(y);
        y_total.append(y_r);
    
    for k in range(np.size(y_total,0)):
        types = [];
        for l in range(np.size(y_total,1)):
            types.append((type(y_total[k][l])));
        matches_str = [1 for x in types if x==str];
        matches_int = [1 for x in types if x==int];
        matches_float = [1 for x in types if x==float];
        if(len(matches_str) == len(types)):
            var_type_metas.append('string');
            var_type_metas_2.append('categorical');
            incompatibilities_metas.append('no');
            ranges_metas.append(list(set(y_total[k]))); incomp_pos_metas.append('-');
        elif(len(matches_int) == len(types)):
            if('year' in metas_features[j])|('Date' in metas_features[j])|('date' in metas_features[j])|('year' in metas_features[j])|('yr' in metas_features[j])|('Dates' in metas_features[j])|('Year' in metas_features[j])|('YEAR' in metas_features[j]):
                var_type_metas.append('date');
            else:
                var_type_metas.append('int');
#            var_type_metas_2.append('numeric');
            incompatibilities_metas.append('no');
            a = np.str(formatNumber(np.nanmin(y_total[k])))+','+np.str(formatNumber(np.nanmax(y_total[k])));
            ranges_metas.append(a.split(',')); incomp_pos_metas.append('-');
            
            #NEW condition
            if(np.nanmin(y_total[k]) == 0)&(np.nanmax(y_total[k]) == 1):
                var_type_metas_2.append('categorical');
            else:
                var_type_metas_2.append('numeric');
            
        elif(len(matches_float) == len(types)):
            if('year' in metas_features[j])|('Date' in metas_features[j])|('date' in metas_features[j])|('year' in metas_features[j])|('yr' in metas_features[j])|('Dates' in metas_features[j])|('Year' in metas_features[j])|('YEAR' in metas_features[j]):
                var_type_metas.append('date');
            else:
                var_type_metas.append('float');
            var_type_metas_2.append('numeric');
            incompatibilities_metas.append('no');
            a = np.str(formatNumber(np.nanmin(y_total[k])))+','+np.str(formatNumber(np.nanmax(y_total[k])));
            ranges_metas.append(a.split(',')); incomp_pos_metas.append('-');
        elif((len(matches_int)+len(matches_float) == len(types))&(len(matches_float)!=0)&(len(matches_int)!=0)):
            var_type_metas.append('float');
            var_type_metas_2.append('numeric');
            incompatibilities_metas.append('no');
            a = np.str(formatNumber(np.nanmin(y_total[k])))+','+np.str(formatNumber(np.nanmax(y_total[k])));
            ranges_metas.append(a.split(',')); incomp_pos_metas.append('-');          
        else:
            var_type_metas.append('unknown');
            var_type_metas_2.append('unknown');
            incompatibilities_metas.append("yes, unknown type of data");
            ranges_metas.append(list(set(y_total[k])));    
            p = [i for i,x in enumerate(types) if ((x==str)&(np.str(y_total[k][i]).strip()!=''))];    
            incomp_pos_metas.append(p);
        
    #missing values
    features_missing_values_metas = [];
    bad_features_metas = [];
    bad_features_ind_metas = [];
    fair_features_metas = [];
    fair_features_ind_metas = [];
    good_features_metas = [];
    good_features_ind_metas = [];
    features_state_metas = [];
    a_total_metas = 0;
        
#        print("Length of pos_metas:", np.str((pos_metas)));
        
    for j in range(len(pos_metas)):
        c1 = 0;
        for index, s in enumerate(y_total[j]):
            if(s == ''):
                c1 = c1+1;
#            c1 = np.around((100*c1)/number_of_rows, 2);        
        features_missing_values_metas.append(c1);
        if(c1>=(number_of_rows/2)):
            bad_features_metas.append(metas_features[j]);
            bad_features_ind_metas.append(j);
            features_state_metas.append('bad');
        elif((c1>0)&(c1<(number_of_rows/2))):
            fair_features_metas.append(metas_features[j]);
            fair_features_ind_metas.append(j);
            features_state_metas.append('fair');          
        elif(c1==0):
            good_features_metas.append(metas_features[j]);
            good_features_ind_metas.append(j);
            features_state_metas.append('good');
        a_total_metas = a_total_metas+c1;
    
    #outlier detection
    outliers_ind_metas = [];
    y_score_metas = [];
    outliers_pos_metas = [];
    if(outlier_detection_method_id == 1):
        for j in range(len(pos_metas)):
            if((var_type_metas[j] == 'string')|(var_type_metas[j] == 'unknown')):
                outliers_ind_metas.append('not-applicable');
                y_score_metas.append('-'); outliers_pos_metas.append('-');
            elif(((var_type_metas[j] == 'float')|(var_type_metas[j] == 'date')|(var_type_metas[j] == 'int'))&(features_missing_values_metas[j]!=number_of_rows)):
                y_f = [];
                for index, s in enumerate(y_total[j]):
                    if((s!='')&(isinstance(s,str)==False)):
                        y_f.append(s);
                    b = np.asarray(y_f);        
                    [z_scores, outliers_ind_z_scores] = outliers_z_score(b);
                    y_score_metas.append(np.mean(z_scores));
                    if(len(outliers_ind_z_scores[0]) != 0):
                        if(var_type_metas[j] == 'int')&(np.max(b)==1)&(np.min(b)==0):
                            outliers_ind_metas.append('no');
                            outliers_pos_metas.append('-');
                        else:
                            outliers_pos_metas.append(list(outliers_ind_z_scores[0]));
                            outliers_ind_metas.append('yes');
                    else:
                        outliers_ind_metas.append('no');
                        outliers_pos_metas.append('-');
            elif(((var_type_metas[j] == 'float')|(var_type_metas[j] == 'date')|(var_type_metas[j] == 'int'))&(features_missing_values_metas[j]==number_of_rows)):
                outliers_ind_metas.append('no');
                outliers_pos_metas.append('-');    
    elif(outlier_detection_method_id == 2):
        for j in range(len(pos_metas)):
            if((var_type_metas[j] == 'string')|(var_type_metas[j] == 'unknown')):
                outliers_ind_metas.append('not-applicable');
                y_score_metas.append('-'); outliers_pos_metas.append('-');
            elif((var_type_metas[j] == 'float')|(var_type_metas[j] == 'date')|(var_type_metas[j] == 'int')):
                y_f = [];
                for index, s in enumerate(y_total[j]):
                    if((s!='')&(isinstance(s,str)==False)):
                        y_f.append(s);
                    b = np.asarray(y_f);
                    [iqr, outliers_ind_iqr] = outliers_iqr(b);
                    y_score_metas.append(np.mean(iqr));
                    if(len(outliers_ind_iqr[0]) != 0):
                        if(var_type_metas[j] == 'int')&(np.max(b)==1)&(np.min(b)==0):
                            outliers_ind_metas.append('no');
                            outliers_pos_metas.append('-');
                        else:
                            outliers_pos_metas.append(list(outliers_ind_iqr[0]));
                            outliers_ind_metas.append('yes');
                    else:
                        outliers_ind_metas.append('no');
                        outliers_pos_metas.append('-');
            elif(((var_type_metas[j] == 'float')|(var_type_metas[j] == 'date')|(var_type_metas[j] == 'int'))&(features_missing_values_metas[j]==number_of_rows)):
                outliers_ind_metas.append('no');
                outliers_pos_metas.append('-');
    elif(outlier_detection_method_id == 3):
        for j in range(len(pos_metas)):
            if((var_type_metas[j] == 'string')|(var_type_metas[j] == 'unknown')):
                outliers_ind_metas.append('not-applicable');
                y_score_metas.append('-'); outliers_pos_metas.append('-');
            elif((var_type_metas[j] == 'float')|(var_type_metas[j] == 'date')|(var_type_metas[j] == 'int')):
                y_f = [];
                for index, s in enumerate(y_total[5]):
                    if((s!='')&(isinstance(s,str)==False)):
                        y_f.append(s);
                    b = np.asarray(y_f);
                    h_min = grubbs.min_test_indices(b);
                    h_max = grubbs.max_test_indices(b);
                    h_tol = np.unique(np.union1d(h_min, h_max));
                    y_score_metas.append(np.mean(grubbs.test(b, alpha=0.05)));
                    if(len(h_tol) != 0):
                        if(var_type_metas[j] == 'int')&(np.max(b)==1)&(np.min(b)==0):
                            outliers_ind_metas.append('no');
                            outliers_pos_metas.append('-');
                        else:
                            outliers_pos_metas.append(list(h_tol));
                            outliers_ind_metas.append('yes');
                    else:
                        outliers_ind_metas.append('no');
                        outliers_pos_metas.append('-');
            elif(((var_type_metas[j] == 'float')|(var_type_metas[j] == 'date')|(var_type_metas[j] == 'int'))&(features_missing_values_metas[j]==number_of_rows)):
                outliers_ind_metas.append('no');
                outliers_pos_metas.append('-');
    elif(outlier_detection_method_id == 4):
        c1 = 0;
        for j in range(len(pos_metas)):
            outliers_ind_metas.append('not-applicable');
            outliers_pos_metas.append('-');
    else:
        for j in range(len(pos_metas)):  
            outliers_pos_metas.append('-');
            outliers_ind_metas.append('no');
                        
    return [features_total, metas_features, pos_metas, y_total, var_type_metas, var_type_metas_2, 
            features_state_metas, incompatibilities_metas, features_missing_values_metas, bad_features_metas, 
            bad_features_ind_metas,fair_features_metas, fair_features_ind_metas, good_features_metas,
            good_features_ind_metas,a_total_metas, outliers_ind_metas, y_score_metas, outliers_pos_metas, ranges_metas, incomp_pos_metas];    


def defineVocabulary():
    wb = open_workbook('pSS_reference_model.xls');
    sheet_names = wb.sheet_names();
    sheet = wb.sheet_by_name(sheet_names[0]);

    harmonicss_ref_dict = sheet.col_values(0,1);
    harmonicss_ref_range = sheet.col_values(3,1);
    harmonicss_ref_class = sheet.col_values(5,1);

    return [harmonicss_ref_dict, harmonicss_ref_range, harmonicss_ref_class];


def defineVocabulary_xml():
    tree = ET.parse('pSS_reference_model.xml');
    root = tree.getroot();

    harmonicss_ref_dict = [];
    harmonicss_ref_range = [];
    harmonicss_ref_class = [];

    for s in range(np.size(root)):
        for e in root[s].iter():
            for j in e:
                harmonicss_ref_dict.append(j.tag.replace("_"," "));
                harmonicss_ref_range.append(j.attrib['range_values']);
                harmonicss_ref_class.append(j.attrib['class_no']);
    
    return [harmonicss_ref_dict, harmonicss_ref_range, harmonicss_ref_class];


def QualityAssessment_S3(data_org, total_features, c, ranges, pos_metas, ranges_metas, total_features_1, total_features_metas):
    #define vocabulary
#    [harmonicss_ref_dict, harmonicss_ref_range, harmonicss_ref_class] = defineVocabulary();
    [harmonicss_ref_dict, harmonicss_ref_range, harmonicss_ref_class] = defineVocabulary_xml();
    
    total_features_clean = [];
    regex = re.compile('\(.+?\)');
    for i in range(c):
        total_features_clean.append(regex.sub('', total_features[i]));
        s = total_features_clean[i];
        if(s[-1] == ' '):
            total_features_clean[i] = s[0:len(s)-1];
    
    total_features_1_clean = [];
    for i in range(np.size(total_features_1)):
        total_features_1_clean.append(regex.sub('', total_features_1[i]));
        s = total_features_1_clean[i];
        if(s[-1] == ' '):
            total_features_1_clean[i] = s[0:len(s)-1];

    total_features_metas_clean = [];
    for i in range(np.size(total_features_metas)):
        total_features_metas_clean.append(regex.sub('', total_features_metas[i]));
        s = total_features_metas_clean[i];
        if(s[-1] == ' '):
            total_features_metas_clean[i] = s[0:len(s)-1];
            
    match_list = [];
    match_class = [];
    list_raw = [];
    list_ref = [];
    list_raw_class = [];
    list_score_class = [];
    list_match_type = [];
    list_range_ref = [];
    count0 = 0;
    
    for s in total_features_clean:
        flag = 0;
        synonyms = [];
        for syn in wn.synsets(s):
            for l in syn.lemmas():
                synonyms.append(l.name());
                
        if(len(synonyms) != 0):
            for ss in synonyms:
                count1 = 0;
                for sss in harmonicss_ref_dict:
                    sim = jaro(ss.lower(), sss.lower());
                    if(sim > 0.9):
#                        print('Found homonymous match between', s, 'and', sss);
                        match_list.append('('+s+','+sss+')');
                        match_class.append('('+s+'->'+harmonicss_ref_class[count1]+')');
                        list_raw.append(total_features[count0]); list_ref.append(sss); list_raw_class.append(harmonicss_ref_class[count1]);
                        list_score_class.append(np.around(sim,2)); list_match_type.append('exact'); flag = 1; break;
                    count1 = count1 + 1;
                    
                if(flag == 1):
                    break;
        else:
            count2 = 0;
            for sss in harmonicss_ref_dict:
                a = s.lower();
                b = sss.lower();
            
                x = intersect(a, b);
                x_t = ''.join(x);
                x_t = ' '.join(x_t.split());
            
                match = SequenceMatcher(None,a,b).find_longest_match(0,len(a),0,len(b));
                sim = jaro(a,b);

                if(((a in b)|(b in a))&('ast' not in a)&('alt' not in a)&('alp' not in a)&('first symptom' not in a)):
#                    print('Found partial match between0', s, 'and', sss);
                    match_list.append('('+s+','+sss+')'); list_match_type.append('partial');
                    match_class.append('('+s+'->'+harmonicss_ref_class[count2]+')');
                    list_raw.append(total_features[count0]); list_ref.append(sss); list_raw_class.append(harmonicss_ref_class[count2]);
                    list_score_class.append('1'); flag = 1; break;
                if(('first' in a)&('visit' in a)):
#                    print('Found partial match between1', s, 'and', 'Age at inclusion');
                    match_list.append('('+s+','+'Age at inclusion'+')'); list_match_type.append('partial');
                    match_class.append('('+s+'->'+'1'+')');
                    list_raw.append(total_features[count0]); list_ref.append('Age at inclusion'); list_raw_class.append('1');
                    list_score_class.append('1'); flag = 2; break;
                elif(('last' in a)&('visit' in a)&('urin' not in a))|(('follow-up' in a)&('during' not in a)):
#                    print('Found partial match between2', s, 'and', 'Age at last follow-up');
                    match_list.append('('+s+','+'Age at last follow-up'+')'); list_match_type.append('partial');
                    match_class.append('('+s+'->'+'1'+')');
                    list_raw.append(total_features[count0]); list_ref.append('Age at last follow-up'); list_raw_class.append('1');
                    list_score_class.append('1'); flag = 2; break;
                elif(('first symptom' in a[match.a:match.a+match.size])&(('year' in a)|('age' in a))):
#                    print('Found partial match between3', s, 'and', sss);
                    match_list.append('('+s+','+sss+')'); list_match_type.append('partial');
                    match_class.append('('+s+'->'+harmonicss_ref_class[count2]+')');
                    list_raw.append(total_features[count0]); list_ref.append(sss); list_raw_class.append(harmonicss_ref_class[count2]);
                    list_score_class.append('1'); flag = 1; break;
                elif(('diagnosis' in a[match.a:match.a+match.size])&(('year' in a)|('age' in a))):
#                    print('Found partial match between4', s, 'and', 'Age at diagnosis of pSS');
                    match_list.append('('+s+','+sss+')');
                    match_class.append('('+s+'->'+harmonicss_ref_class[count2]+')'); list_match_type.append('partial');
                    list_raw.append(total_features[count0]); list_ref.append(sss); list_raw_class.append(harmonicss_ref_class[count2]);
                    list_score_class.append('1'); flag = 1; break;
                elif('whole saliva' in a[match.a:match.a+match.size]):
#                    print('Found partial match between5', s, 'and', sss);
                    match_list.append('('+s+','+sss+')');
                    match_class.append('('+s+'->'+harmonicss_ref_class[count2]+')'); list_match_type.append('partial');
                    list_raw.append(total_features[count0]); list_ref.append(sss); list_raw_class.append(harmonicss_ref_class[count2]);
                    list_score_class.append('1'); flag = 2; break;
                elif('ocular stain' in a[match.a:match.a+match.size]):
#                    print('Found partial match between6', s, 'and', sss);
                    match_list.append('('+s+','+sss+')');
                    match_class.append('('+s+'->'+harmonicss_ref_class[count2]+')'); list_match_type.append('partial');
                    list_raw.append(total_features[count0]); list_ref.append(sss); list_raw_class.append(harmonicss_ref_class[count2]);
                    list_score_class.append('1'); flag = 1; break;
                elif('rf' in a)&('perf' not in a)&(b == 'rheumatoid factor'):
#                    print('Found partial match between7', s, 'and', sss);
                    match_list.append('('+s+','+sss+')');
                    match_class.append('('+s+'->'+harmonicss_ref_class[count2]+')'); list_match_type.append('partial');
                    list_raw.append(total_features[count0]); list_ref.append(sss); list_raw_class.append(harmonicss_ref_class[count2]);
                    list_score_class.append('1'); flag = 3; break;
                elif('msg' in a)&(b == 'minor salivary gland biopsy'):
#                    print('Found partial match between8', s, 'and', sss);
                    match_list.append('('+s+','+sss+')');
                    match_class.append('('+s+'->'+harmonicss_ref_class[count2]+')'); list_match_type.append('partial');
                    list_raw.append(total_features[count0]); list_ref.append(sss); list_raw_class.append(harmonicss_ref_class[count2]);
                    list_score_class.append('1'); flag = 2; break;
                elif('monoclonal' in a[match.a:match.a+match.size]):
#                    print('Found partial match between10', s, 'and', sss);
                    match_list.append('('+s+','+sss+')');
                    match_class.append('('+s+'->'+harmonicss_ref_class[count2]+')'); list_match_type.append('partial');
                    list_raw.append(total_features[count0]); list_ref.append(sss); list_raw_class.append(harmonicss_ref_class[count2]);
                    list_score_class.append('1'); flag = 1; break;
                elif('urin' in a[match.a:match.a+match.size])&('during' not in a):
#                    print('Found partial match between11', s, 'and', sss);
                    match_list.append('('+s+','+sss+')');
                    match_class.append('('+s+'->'+harmonicss_ref_class[count2]+')'); list_match_type.append('partial');
                    list_raw.append(total_features[count0]); list_ref.append(sss); list_raw_class.append(harmonicss_ref_class[count2]);
                    list_score_class.append('1'); flag = 1; break;
                elif('vasculiti' in a[match.a:match.a+match.size])&('pns' not in a)&('pns' not in b):
#                    print('Found partial match between12', s, 'and', sss);
                    match_list.append('('+s+','+sss+')');
                    match_class.append('('+s+'->'+harmonicss_ref_class[count2]+')'); list_match_type.append('partial');
                    list_raw.append(total_features[count0]); list_ref.append(sss); list_raw_class.append(harmonicss_ref_class[count2]);
                    list_score_class.append('1'); flag = 1; break;
                elif('cryo' in a[match.a:match.a+match.size])&('vasculiti' not in a[match.a:match.a+match.size]):
#                    print('Found partial match between13', s, 'and', sss);
                    match_list.append('('+s+','+sss+')'); list_match_type.append('partial');
                    match_class.append('('+s+'->'+harmonicss_ref_class[count2]+')');
                    list_raw.append(total_features[count0]); list_ref.append(sss); list_raw_class.append(harmonicss_ref_class[count2]);
                    list_score_class.append('1'); flag = 1; break;
                elif('sgot' in a)&(b == 'ast'):
#                    print('Found partial match between13', s, 'and', sss);
                    match_list.append('('+s+','+sss+')'); list_match_type.append('partial');
                    match_class.append('('+s+'->'+harmonicss_ref_class[count2]+')');
                    list_raw.append(total_features[count0]); list_ref.append(sss); list_raw_class.append(harmonicss_ref_class[count2]);
                    list_score_class.append('1'); flag = 1; break;
                elif('alp' in a)&(b == 'alp')&('palpable' not in a):
#                    print('Found partial match between13', s, 'and', sss);
                    match_list.append('('+s+','+sss+')'); list_match_type.append('partial');
                    match_class.append('('+s+'->'+harmonicss_ref_class[count2]+')');
                    list_raw.append(total_features[count0]); list_ref.append(sss); list_raw_class.append(harmonicss_ref_class[count2]);
                    list_score_class.append('1'); flag = 1; break;
                elif('sgpt' in a)&(b == 'alt'):
#                    print('Found partial match between13', s, 'and', sss);
                    match_list.append('('+s+','+sss+')'); list_match_type.append('partial');
                    match_class.append('('+s+'->'+harmonicss_ref_class[count2]+')');
                    list_raw.append(total_features[count0]); list_ref.append(sss); list_raw_class.append(harmonicss_ref_class[count2]);
                    list_score_class.append('1'); flag = 1; break;
                elif('hirmer' in a[match.a:match.a+match.size]):
#                    print('Found partial match between13', s, 'and', sss);
                    match_list.append('('+s+','+sss+')'); list_match_type.append('partial');
                    match_class.append('('+s+'->'+harmonicss_ref_class[count2]+')');
                    list_raw.append(total_features[count0]); list_ref.append(sss); list_raw_class.append(harmonicss_ref_class[count2]);
                    list_score_class.append('1'); flag = 1; break;             
                elif('dyspareunia' in a[match.a:match.a+match.size]):
#                    print('Found partial match between14', s, 'and', sss);
                    match_list.append('('+s+','+sss+')'); list_match_type.append('partial');
                    match_class.append('('+s+'->'+harmonicss_ref_class[count2]+')');
                    list_raw.append(total_features[count0]); list_ref.append(sss); list_raw_class.append(harmonicss_ref_class[count2]);
                    list_score_class.append('1'); flag = 1; break;               
                elif('fatigue' in a[match.a:match.a+match.size]):
#                    print('Found partial match between15', s, 'and', sss);
                    match_list.append('('+s+','+sss+')'); list_match_type.append('partial');
                    match_class.append('('+s+'->'+harmonicss_ref_class[count2]+')');
                    list_raw.append(total_features[count0]); list_ref.append(sss); list_raw_class.append(harmonicss_ref_class[count2]);
                    list_score_class.append('1'); flag = 1; break;              
                elif(('cns' in a[match.a:match.a+match.size])|('central neuro' in a[match.a:match.a+match.size])):
#                    print('Found partial match between16', s, 'and', sss);
                    match_list.append('('+s+','+sss+')'); list_match_type.append('partial');
                    match_class.append('('+s+'->'+harmonicss_ref_class[count2]+')');
                    list_raw.append(total_features[count0]); list_ref.append(sss); list_raw_class.append(harmonicss_ref_class[count2]);
                    list_score_class.append('1'); flag = 1; break;           
                elif(('pns' in a[match.a:match.a+match.size])|('peripheral neuro' in a[match.a:match.a+match.size])):
#                    print('Found partial match between17', s, 'and', sss);
                    match_list.append('('+s+','+sss+')'); list_match_type.append('partial');
                    match_class.append('('+s+'->'+harmonicss_ref_class[count2]+')');
                    list_raw.append(total_features[count0]); list_ref.append(sss); list_raw_class.append(harmonicss_ref_class[count2]);
                    list_score_class.append('1'); flag = 1; break;        
                elif(('renal' in a[match.a:match.a+match.size])|('renal' in b[match.a:match.a+match.size])):
#                    print('Found partial match between18', s, 'and', sss);
                    match_list.append('('+s+','+sss+')'); list_match_type.append('partial');
                    match_class.append('('+s+'->'+harmonicss_ref_class[count2]+')');
                    list_raw.append(total_features[count0]); list_ref.append(sss); list_raw_class.append(harmonicss_ref_class[count2]);
                    list_score_class.append('1'); flag = 1; break;        
                elif('urine' in a[match.a:match.a+match.size]):
#                    print('Found partial match between19', s, 'and', sss);
                    match_list.append('('+s+','+sss+')'); list_match_type.append('partial');
                    match_class.append('('+s+'->'+harmonicss_ref_class[count2]+')');
                    list_raw.append(total_features[count0]); list_ref.append(sss); list_raw_class.append(harmonicss_ref_class[count2]);
                    list_score_class.append('1'); flag = 1; break;
                elif('crp' in a)&(b == 'increased c-reactive protein'):
#                    print('Found partial match between19', s, 'and', sss);
                    match_list.append('('+s+','+sss+')'); list_match_type.append('partial');
                    match_class.append('('+s+'->'+harmonicss_ref_class[count2]+')');
                    list_raw.append(total_features[count0]); list_ref.append(sss); list_raw_class.append(harmonicss_ref_class[count2]);
                    list_score_class.append('1'); flag = 3; break;
                elif('globulins' in a[match.a:match.a+match.size]):
#                    print('Found partial match between20', s, 'and', sss);
                    match_list.append('('+s+','+sss+')'); list_match_type.append('partial');
                    match_class.append('('+s+'->'+harmonicss_ref_class[count2]+')');
                    list_raw.append(total_features[count0]); list_ref.append(sss); list_raw_class.append(harmonicss_ref_class[count2]);
                    list_score_class.append('1'); flag = 1; break; 
                elif(('lymphoma' in a[match.a:match.a+match.size])|('lymphadenopathy' in a[match.a:match.a+match.size])):
#                    print('Found partial match between21', s, 'and', sss);
                    match_list.append('('+s+','+sss+')'); list_match_type.append('partial');
                    match_class.append('('+s+'->'+harmonicss_ref_class[count2]+')');
                    list_raw.append(total_features[count0]); list_ref.append(sss); list_raw_class.append(harmonicss_ref_class[count2]);
                    list_score_class.append('1'); flag = 1; break;
                elif(('dry' in a)&('mouth' in a)&(b == 'oral dryness')):
#                    print('Found partial match between22', s, 'and', sss);
                    match_list.append('('+s+','+sss+')'); list_match_type.append('partial');
                    match_class.append('('+s+'->'+harmonicss_ref_class[count2]+')');
                    list_raw.append(total_features[count0]); list_ref.append(sss); list_raw_class.append(harmonicss_ref_class[count2]); 
                    list_score_class.append('1'); flag = 3; break;
                elif(('dry' in a)&('eyes' in a)&(b == 'ocular dryness')):
#                    print('Found partial match between23', s, 'and', sss);
                    match_list.append('('+s+','+sss+')'); list_match_type.append('partial');
                    match_class.append('('+s+'->'+harmonicss_ref_class[count2]+')');
                    list_raw.append(total_features[count0]); list_ref.append(sss); list_raw_class.append(harmonicss_ref_class[count2]); 
                    list_score_class.append('1'); flag = 3; break;
                elif(('haematologic' in a[match.a:match.a+match.size])|('hematologic' in a[match.a:match.a+match.size])):
#                    print('Found partial match between24', s, 'and', sss);
                    match_list.append('('+s+','+sss+')'); list_match_type.append('partial');
                    match_class.append('('+s+'->'+harmonicss_ref_class[count2]+')');
                    list_raw.append(total_features[count0]); list_ref.append(sss); list_raw_class.append(harmonicss_ref_class[count2]);
                    list_score_class.append('1'); flag = 1; break;
                elif((a in 'hgb')&(b in 'haematological domain')):
#                    print('Found partial match between25', s, 'and', sss);
                    match_list.append('('+s+','+sss+')'); list_match_type.append('partial');
                    match_class.append('('+s+'->'+harmonicss_ref_class[count2]+')');
                    list_raw.append(total_features[count0]); list_ref.append(sss); list_raw_class.append(harmonicss_ref_class[count2]);
                    list_score_class.append('1'); flag = 1; break;
                elif(sim>=0.9):
#                    print('Found exact match between26', s, 'and', sss);
                    match_list.append('('+s+','+sss+')'); list_match_type.append('exact');
                    match_class.append('('+s+'->'+harmonicss_ref_class[count2]+')');
                    list_raw.append(total_features[count0]); list_ref.append(sss); list_raw_class.append(harmonicss_ref_class[count2]);
                    list_score_class.append(np.around(sim, 2)); flag = 1; break;
                count2 = count2 + 1;
                
        #update the final range
        if(flag == 1)|(flag == 2):
            p = [i for i,x in enumerate(harmonicss_ref_dict) if x == sss];
            
            if(harmonicss_ref_range[p[0]] != '-'):
                list_range_ref.append(harmonicss_ref_range[p[0]]);
            else:
                if(count0 < np.size(total_features_1)):
                    q = [i for i,x in enumerate(total_features_1_clean) if x == s];
#                    print(s); print(sss); print(q); print('');
                    ranges_f = str(ranges[q[0]]).replace("'",'');
                    list_range_ref.append(ranges_f);
                else:
                    q = [i for i,x in enumerate(total_features_metas_clean) if x == s];
#                    print("Metas"); print(s); print(sss); print(q); print('');
                    ranges_f = str(ranges_metas[q[0]]).replace("'",'');
                    list_range_ref.append(ranges_f);
        elif(flag == 3):
            list_range_ref.append('[yes,no]');
                
        count0 = count0 + 1;
      
    return [list_raw, list_raw_class, list_ref, list_score_class, list_match_type, list_range_ref];


def similarity_detection(wb):
    sheet_names = wb.sheet_names();
    xl = wb.sheet_by_name(sheet_names[0]);
    ncols = xl.ncols;
    cmatrix = np.zeros((ncols,ncols));
    pmatrix = np.zeros((ncols,ncols));
    
    features_total = [];
    for col in range(ncols):
        features_total.append((xl.cell(0,col).value));
    
    for j in range(ncols):
        a = xl.col_values(j,1);
        for k in range(j+1,ncols):
            b = xl.col_values(k,1);
            try:
                a = pd.to_numeric(a, errors='coerce');
                b = pd.to_numeric(b, errors='coerce');
                if(np.sum(np.isnan(a)) < 20)&(np.sum(np.isnan(b)) < 20):
                    [c,p] = spearmanr(a,b,nan_policy='omit');
                else:
                    c = 0; p = 1;
            except:
                c = 0; p = 1;
            cmatrix[j,k] = c;
            pmatrix[j,k] = p;
    cmatrix = cmatrix + cmatrix.T;
    pmatrix = pmatrix + pmatrix.T;
    np.fill_diagonal(cmatrix,1);
    np.fill_diagonal(pmatrix,0);
    fnans = np.isnan(cmatrix);
    pnans = np.isnan(pmatrix);
    cmatrix[fnans] = 0;
    pmatrix[pnans] = 100;
    f_cmatrix = np.where((np.tril(cmatrix,-1) >= 0.9) & (np.tril(cmatrix,-1) <= 1));

#    cmatrix[cmatrix<0] = 0;
##    sns.set(font_scale = 1);
#    sns_plot = sns.heatmap(cmatrix);
##    sns_plot.set_title('Correlation matrix for the UoA dataset');
#    sns_plot.set_xlabel('Features (n)');
#    sns_plot.set_ylabel('Features (n)');
#    plt.tight_layout();
#    fig = sns_plot.figure;
#    fig.savefig("curation/v2/data_curation_corrcoef_UoA.png");
#    fig.clf(); plt.close(fig);
    
    f_cmatrix_names = [];
    r_values = [];
    p_values = [];
    for k in range(np.size(f_cmatrix,1)):
        f_cmatrix_names.append('('+features_total[f_cmatrix[0][k]]+','+features_total[f_cmatrix[1][k]]+')');
        r_values.append(cmatrix[f_cmatrix[0][k], f_cmatrix[1][k]]);
        p_values.append(pmatrix[f_cmatrix[0][k], f_cmatrix[1][k]]);
    
    total_features_clean = [];
    regex = re.compile('\(.+?\)');
    for i in range(ncols):
        total_features_clean.append(regex.sub('', features_total[i]));
        s = total_features_clean[i];
        if(s[-1] == ' '):
            total_features_clean[i] = s[0:len(s)-1];
    
    jmatrix = np.zeros((ncols,ncols));
    for m in range(ncols):
        for n in range(m+1,ncols):
            jdist = jaro(total_features_clean[m], total_features_clean[n]);
            jmatrix[m,n] = jdist;
    
    jmatrix = jmatrix + jmatrix.T;       
    f_jmatrix = np.where((np.tril(jmatrix,-1) >= 0.97));
    f_jmatrix_names = [];
    j_values = [];
    for k in range(np.size(f_jmatrix,1)):
        f_jmatrix_names.append('('+features_total[f_jmatrix[0][k]]+','+features_total[f_jmatrix[1][k]]+')');
        j_values.append(jmatrix[f_jmatrix[0][k], f_jmatrix[1][k]]);

##    sns.set(font_scale = 1.2);
#    sns_plot = sns.heatmap(jmatrix);
##    sns_plot.set_title('Lexical distance matrix for the UoA dataset');
#    sns_plot.set_xlabel('Features (n)');
#    sns_plot.set_ylabel('Features (n)');
#    plt.tight_layout();
#    fig = sns_plot.figure;
#    fig.savefig("curation/v2/data_curation_distance_UoA.png");
#    fig.clf(); plt.close(fig);
    
#    for j in range(np.size(f_cmatrix_names,0)):
#        print(f_cmatrix_names[j]);
#        print(r_values[j]);
#        print(p_values[j]);
#        print("");
    
    return [cmatrix, f_cmatrix, pmatrix, f_cmatrix_names, r_values, p_values, jmatrix, f_jmatrix, f_jmatrix_names, j_values, features_total];


@app.route("/features/", methods = ['GET', 'POST'])
def get_feature_values():
    path = 'data/UoA_small/demo-pSS-code-v3.xlsx';
    wb = open_workbook(path);
    sheet_names = wb.sheet_names();
    xl = wb.sheet_by_name(sheet_names[0]);
    f = xl.row_values(0);
    f = [x.replace('\n','') for x in f];
    
    feature_id = request.args.get('feature_id', type=int);
    if(feature_id > np.size(f))|(feature_id <= 0):
        return json.dumps({'Error':'Selected feature is out of bounds!'});
    
    b0 = xl.col_values(feature_id-1, 1);
#    b = list(filter(None, b0));
    sel_f = f[feature_id-1];
    sel_f_values = np.str(b0).replace('\n','');
    
    
    b1 = xl.col_values(np.size(f)-1, 1);
    sel_t_values = np.str(b1).replace('\n','');
    dictionary = [{'Selected feature':sel_f,
                   'Selected feature values':sel_f_values,
                   'Target values':sel_t_values}];
    
    d = create_wr_io('results_feature.txt', dictionary);
    
    return jsonify(d);


@app.route("/features/names/", methods = ['GET', 'POST'])
def get_features_names():
    path = 'data/UoA_small/demo-pSS-code-v3.xlsx';
    wb = open_workbook(path);
    sheet_names = wb.sheet_names();
    xl = wb.sheet_by_name(sheet_names[0]);
    f = xl.row_values(0);
    f = [x.replace('\n','') for x in f];

    dictionary = [{'Features': f}];
    
    return jsonify(dictionary);


def jsonify_data_curator(path, imputation_method_id, outlier_detection_method_id, descriptive_feature_id, wb2, bad_features_ind, bad_features_ind_metas, metas_features, cmatrix, f_cmatrix):
    outlier_detection_methods = ['z-score', 'IQR', 'Grubbs', 'None'];
    imputation_methods = ['mean', 'random', 'None'];
    
    if(imputation_method_id == None):
        imputation_method_id = 3;
    
    if(outlier_detection_method_id == None):
        outlier_detection_method_id = 4;
        
    wb = open_workbook(path);
    sheet_names = wb.sheet_names();
    xl = wb.sheet_by_name(sheet_names[0]);
    
    nof = xl.cell(1,4).value;
    noi = xl.cell(2,4).value;
    df = xl.cell(3,4).value;
    cf = xl.cell(4,4).value;
    un = xl.cell(5,4).value;
    mv = xl.cell(6,4).value;
    f = xl.col_values(0, 10);
    f = [f[i].replace('\n','')  for i,x in enumerate(f)];
    
    ranges = xl.col_values(5, 10);
    t = xl.col_values(9, 10);
    t2 = xl.col_values(11, 10);
    mvf = xl.col_values(13, 10);
    s = xl.col_values(16, 10);
    o = xl.col_values(18, 10);
    inco = xl.col_values(20, 10);
    
#    wb2 = open_workbook(path);
    sheet_names2 = wb2.sheet_names();
    xl2 = wb2.sheet_by_name(sheet_names2[0]);
    
    if(descriptive_feature_id != 0):
        m = [];
        med = [];
        std = [];
        sk = [];
        kurt = [];
        
        b0 = xl2.col_values(descriptive_feature_id-1, 1);
        b = list(filter(None, b0));
        
        try:
            m.append(np.around(np.mean(b), 2));
        except:
            m.append('None');

        try:
            med.append(np.around(np.median(b), 2));
        except:
            med.append('None');

        try:
            std.append(np.around(np.std(b), 2));
        except:
            std.append('None');

        try:
            sk.append(np.around(scipy.stats.skew(b), 2));
        except:
            sk.append('None');

        try:
            kurt.append(np.around(scipy.stats.kurtosis(b), 2));
        except:
            kurt.append('None');
    
    bf_ind = [i for i,x in enumerate(s) if x=='bad'];
    bf_names = [f[i] for i in np.asarray(bf_ind,int)];
    
#    with Capturing() as output:
#        for i in range(np.size(f_cmatrix,1)):
#            print("(", repr(f[f_cmatrix[0][i]].replace('\n',' ')), ",", 
#                    repr(f[f_cmatrix[1][i]].replace('\n',' ')), ",", 
#                    repr(np.str(np.around(cmatrix[f_cmatrix[0][i], f_cmatrix[1][i]], 2))), ")");
    
#    with Capturing() as output:
    output = [];
    for i in range(np.size(f_cmatrix,1)):
        output.append(["(" + repr(f[f_cmatrix[0][i]].replace('\n',' ')) + "," + 
                    repr(f[f_cmatrix[1][i]].replace('\n',' ')) + "," + 
                    repr(np.str(np.around(cmatrix[f_cmatrix[0][i], f_cmatrix[1][i]], 2))) + ")"]);
    
    python_dict = [{'Number of feature(s)':np.str(nof),
                    'Number of instance(s)':np.str(noi),
                    'Discrete feature(s)':np.str(df),
                    'Continuous feature(s)':np.str(cf),
                    'Unknown feature(s)':np.str(un),
                    'Missing values':np.str(mv),
                    'Selected feature':f[descriptive_feature_id-1],
                    'Selected feature values':[np.str(e).replace('\n','') for e in b0],
                    'Mean':m,
                    'Median':med,
                    'Std':std,
                    'Skewness':sk,
                    'Kurtosis':kurt,
                    'Features':f,
                    'Values':ranges,
                    'Type':t,
                    'Type_2':t2,
                    'Missing values (per feature)':[np.str(e) for e in mvf],
                    'State':s,
                    'Outlier detection method':outlier_detection_methods[outlier_detection_method_id-1],
                    'Outliers':o,
                    'Compatibility issues':inco,
                    'Features with > 50% missing values':bf_names,
                    'Imputation method':imputation_methods[imputation_method_id-1],
#                    'Features with detected outliers':np.str(outliers_ind).replace('\n',' '),          
#                    'Number of outliers per feature with detected outliers':np.str(outliers_num).replace('\n',' '),
#                    'Position of outliers per feature with detected outliers':np.str(outliers_pos),
                    'Highly correlated pair(s) of features':[str(x).replace('\"', '"') for e in output for x in e],
                    'Meta-attribute(s)':metas_features,
                    }];
    
    d = create_wr_io('curation/json_results.txt', python_dict);
    return d;


def data_annotation(data_org):
    c = np.size(data_org,1);
    r = np.size(data_org,0);
    var_type_final = [];
    features = [];
    ranges = [];
    var_type_final_2 = [];
    incompatibilities = [];
    incomp_pos = [];

    for j in range(c):
        features.append(data_org.domain.attributes[j].name);
        if(data_org.domain.attributes[j].is_discrete == True):
            y = data_org.domain.attributes[j].values;
#            var_type_final_2.append('categorical');
            var_type = [];
            y_total = np.zeros(len(y));
            for i in range(len(y)):
                try:
                    y_mod = float(y[i]);
                    [y_mod, flag] = formatNumber_v2(y_mod);
                    y_total[i] = y_mod;
                    if(flag == 1):
                        var_type.append('int');
                    elif(flag == 0):
                        var_type.append('float');                       
                except:
                    y_mod = y[i];
                    if(isinstance(y_mod,str) == True):
                        var_type.append('string');
                    else:
                        var_type.append('unknown');
                    
            matches_str = [1 for x in var_type if x=='string'];
            matches_int = [1 for x in var_type if x=='int'];
            matches_float = [1 for x in var_type if x=='float'];
            
            if(len(matches_str) == len(var_type)):
                var_type_final.append('string'); var_type_final_2.append('categorical');
                ranges.append(y); incompatibilities.append('no'); incomp_pos.append('-');
            elif(len(matches_int) == len(var_type)):
                if('year' in features[j])|('Date' in features[j])|('date' in features[j])|('year' in features[j])|('yr' in features[j])|('Dates' in features[j])|('Year' in features[j])|('YEAR' in features[j]):
                    var_type_final.append('date'); var_type_final_2.append('numeric');
                    a = np.str(formatNumber(np.nanmin(y_total)))+','+np.str(formatNumber(np.nanmax(y_total)));
                    ranges.append(a.split(',')); incompatibilities.append('no');
                else:
                    var_type_final.append('int');
                    a = np.str(formatNumber(np.nanmin(y_total)))+','+np.str(formatNumber(np.nanmax(y_total)));
                    ranges.append(a.split(',')); incompatibilities.append('no'); 
                    
                    #NEW condition
                    if(np.nanmin(y_total) == 0)&(np.nanmax(y_total) == 1):
                        var_type_final_2.append('categorical');
                    else:
                        var_type_final_2.append('numeric');
                incomp_pos.append('-');
            elif((len(matches_int)+len(matches_float) == len(var_type))&(len(matches_float)!=0)&(len(matches_int)!=0)):
                if('year' in features[j])|('Date' in features[j])|('date' in features[j])|('year' in features[j])|('yr' in features[j])|('Dates' in features[j])|('Year' in features[j])|('YEAR' in features[j]):
                    var_type_final.append('date'); var_type_final_2.append('numeric');
                    a = np.str(formatNumber(np.nanmin(y_total)))+','+np.str(formatNumber(np.nanmax(y_total)));
                    ranges.append(a.split(',')); incompatibilities.append('no');
                else:
                    var_type_final.append('float'); var_type_final_2.append('numeric');
                    a = np.str(formatNumber(np.nanmin(y_total)))+','+np.str(formatNumber(np.nanmax(y_total)));
                    ranges.append(a.split(',')); incompatibilities.append('no'); 
                incomp_pos.append('-');             
            elif(len(matches_float) == len(var_type)):
                if('year' in features[j])|('Date' in features[j])|('date' in features[j])|('year' in features[j])|('yr' in features[j])|('Dates' in features[j])|('Year' in features[j])|('YEAR' in features[j]):
                    var_type_final.append('date'); var_type_final_2.append('numeric');
                    a = np.str(formatNumber(np.nanmin(y_total)))+','+np.str(formatNumber(np.nanmax(y_total)));
                    ranges.append(a.split(',')); incompatibilities.append('no');
                else:
                    var_type_final.append('float'); var_type_final_2.append('numeric');
                    a = np.str(formatNumber(np.nanmin(y_total)))+','+np.str(formatNumber(np.nanmax(y_total)));
                    ranges.append(a.split(',')); incompatibilities.append('no');
                incomp_pos.append('-');
            else:
                var_type_final.append('unknown'); var_type_final_2.append('unknown');
                ranges.append(y);
                p = [i for i,x in enumerate(var_type) if ((x=='string')|(x=='unknown'))];
                incomp_names = [y[k] for k in p];
                incomp_names_ind = [];
                for m in range(len(incomp_names)):
                    incomp_names_ind.append([i for i,x in enumerate(data_org[:,j]) if x.list[0]==incomp_names[m]]);
                incomp_names_ind = list(itertools.chain.from_iterable(incomp_names_ind));
                incomp_pos.append(incomp_names_ind);
                incompatibilities.append("yes, unknown type of data");
        else:
            var_type_final_2.append('numeric');
            if('year' in features[j])|('Date' in features[j])|('date' in features[j])|('year' in features[j])|('yr' in features[j])|('Dates' in features[j])|('Year' in features[j])|('YEAR' in features[j]):
                var_type_final.append('date');
                a = np.str(formatNumber(np.nanmin(data_org.X[:,j])))+','+np.str(formatNumber(np.nanmax(data_org.X[:,j])));
                ranges.append(a.split(',')); incompatibilities.append('no'); incomp_pos.append('-');
            else:
                g = np.array(list(data_org.X[:,j]));
                g = g[np.isnan(g) == False];
                if(np.all(g%1==0)):
                    var_type_final.append('int');
                    a = np.str(formatNumber(np.nanmin(data_org.X[:,j])))+','+np.str(formatNumber(np.nanmax(data_org.X[:,j])));
                    ranges.append(a.split(',')); incompatibilities.append('no'); incomp_pos.append('-'); 
                else:
                    var_type_final.append('float');
                    a = np.str(formatNumber(np.nanmin(data_org.X[:,j])))+','+np.str(formatNumber(np.nanmax(data_org.X[:,j])));
                    ranges.append(a.split(',')); incompatibilities.append('no'); incomp_pos.append('-');
                        
    return [r, c, var_type_final, var_type_final_2, ranges, incompatibilities, incomp_pos];


def plot_curator_results(data_org, cmatrix, outliers_ind, good_features_ind):
#    data_org = Orange.data.Table(path);
    imputer = Impute(method=Average());
    my_data = imputer(data_org);
    
    X = my_data.X;
    z = X[:,135];
    z[z>30] = np.mean(z);
    X[:,135] = z;
    cmatrix = np.nan_to_num(cmatrix);
    cmatrix[cmatrix<0] = 0;
    sns.set(font_scale=1.2);
    sns_plot = sns.heatmap(cmatrix);
    sns_plot.set_title('Similarity matrix');
    sns_plot.set_xlabel('Features (n)');
    sns_plot.set_ylabel('Features (n)');
    plt.tight_layout();
    fig = sns_plot.figure;
    fig.savefig("curation/v2/data_curation_corrcoef.png");
    fig.clf(); plt.close(fig);
    
    x_ind = [24, 41, 119, 135];
    
    print(np.str(np.min(X[:,135])));
    print(np.str(np.max(X[:,135])));
    print(np.str(np.mean(X[:,135])));
    
    B = X[:,x_ind];
    sns.set_style("whitegrid");
    sns.set(font_scale=1.2);
    sns_plot = sns.boxplot(data=B, orient="v", width = 0.3);
    sns_plot.set_xticklabels(["Tarpley", "Lymphoma"+"\n"+"score", "URINE PH AT"+"\n"+"LAST VISIT", "HGB"+"\n"+"(Absolute number)"]);
    sns_plot.set_ylabel('value');
    plt.tight_layout();
    fig = sns_plot.figure;
    fig.savefig("curation/v2/IQR.png");
    fig.clf(); plt.close(fig);
    
    #23, 37, 91, 105
    for j in range(len(x_ind)):
        b = X[:,x_ind[j]];
        z_scores = [(b - np.mean(b)) / np.std(b)];
        sns.set(font_scale=1.5);
        sns_plot = sns.distplot(z_scores, bins=10);
        plt.axvline(x=3,color='r',linestyle='--');
        plt.axvline(x=-3,color='r',linestyle='--');
        sns_plot.set_xlabel('z-scores');
        sns_plot.set_ylabel('density');
        fig = sns_plot.figure;
        path = "curation/v2/zscore_dist_"+np.str(j)+".png";
        fig.savefig(path);
        fig.clf(); plt.close(fig);

    X_c = data_org.X[:,good_features_ind];
    clusterer = hdbscan.HDBSCAN(min_cluster_size=2).fit(X_c);
    y_pred_lof = clusterer.outlier_scores_;
    
    sns.set(font_scale=1);
    sns_plot = sns.distplot(y_pred_lof, bins=10, rug=True);
    sns_plot.set_xlabel('LOF values');
    sns_plot.set_ylabel('density');
    fig = sns_plot.figure;
    fig.savefig("curation/v2/LOF_dist.png");
    fig.clf(); plt.close(fig);  


def data_curator(path, imputation_method_id, outlier_detection_method_id):
    print('Imputation method id:', np.str(imputation_method_id));
    print('Outlier detection method id:', np.str(outlier_detection_method_id));
    print(path);
    #construct paths
    fname = path.split('/')[2];
    fname_c = fname.split('.')[0]+"_curated_dataset.xls";
    fname_e = fname.split('.')[0]+"_evaluation_report.xls";
    fname_s = fname.split('.')[0]+"_standardization_report.xls";
    
    path_f_c = 'results/'+fname_c;
    path_f_e = 'results/'+fname_e;
    path_f_s = 'results/'+fname_s;
    
    print(path_f_c);
    print(path_f_e);
    print(path_f_s);
    
    print("Reading data...");
    start = timeit.default_timer();
    wb = open_workbook(path);
    data_org = Orange.data.Table(path);
    stop = timeit.default_timer();
    print('Time: ', np.around(stop - start, 3), 'sec'); print();
    
    imputation_method_id = request.args.get('imputation_method_id', type=int);
    outlier_detection_method_id = request.args.get('outlier_detection_method_id', type=int);
    descriptive_feature_id = request.args.get('descriptive_feature_id', type=int);
    
    if(imputation_method_id is None):
        imputation_method_id = 0;
        
    if(outlier_detection_method_id is None):
        outlier_detection_method_id = 1;

    if(descriptive_feature_id is None):
        descriptive_feature_id = 1;

    np.warnings.filterwarnings('ignore');

    print("Annotating data...");
    start = timeit.default_timer();
    [r, c, var_type_final, var_type_final_2, ranges, incompatibilities, incomp_pos] = data_annotation(data_org);
    stop = timeit.default_timer();
    print('Time: ', np.around(stop - start, 3), 'sec'); print();
    
    print("Assessing the quality of the data...");
    start = timeit.default_timer();
    [features_missing_values, bad_features, bad_features_ind, fair_features, fair_features_ind, 
     good_features, good_features_ind, features_state, a, a_total] = QualityAssessment_S1(data_org, c);
    stop = timeit.default_timer();
    print('Time: ', np.around(stop - start, 3), 'sec'); print();
    
    print("Checking for outliers...");
    start = timeit.default_timer();    
    [outliers_ind, y_score, outliers_pos] = outliers_detection(data_org, c, var_type_final, outlier_detection_method_id, bad_features_ind, fair_features_ind, good_features_ind, features_missing_values);
    stop = timeit.default_timer();
    print('Time: ', np.around(stop - start, 3), 'sec'); print();
            
    print("Handling meta-attribute(s)...");
    start = timeit.default_timer();            
    [features_total, metas_features, pos_metas, y_total, var_type_metas, var_type_metas_2, features_state_metas, incompatibilities_metas, 
     features_missing_values_metas, bad_features_metas, bad_features_ind_metas, fair_features_metas, fair_features_ind_metas, good_features_metas,good_features_ind_metas,
     a_total_metas, outliers_ind_metas, y_score_metas, outliers_pos_metas, ranges_metas, incomp_pos_metas] = metas_handler(data_org, wb, outlier_detection_method_id);
    stop = timeit.default_timer();
    print('Time: ', np.around(stop - start, 3), 'sec'); print();
     
    print("Applying similarity detection...");
    start = timeit.default_timer();    
    np.seterr(divide='ignore', invalid='ignore'); 
    [cmatrix, f_cmatrix, _, _, _, _, _, _, _, _, _] = similarity_detection(wb);
    stop = timeit.default_timer();
    print('Time: ', np.around(stop - start, 3), 'sec'); print();
     
    print("Applying standardization...");
    start = timeit.default_timer();    
    c1 = np.size(data_org.domain.metas);
    total_features = [];
    total_features_1 = [];
    total_features_metas = [];
    for i in range(c):
        total_features.append(data_org.domain.attributes[i].name.replace('\n',' ').replace('  ', ' '));
        total_features_1.append(data_org.domain.attributes[i].name.replace('\n',' ').replace('  ', ' '));
        
    for i in range(c1):
        total_features.append(data_org.domain.metas[i].name.replace('\n',' ').replace('  ', ' '));
        total_features_metas.append(data_org.domain.metas[i].name.replace('\n',' ').replace('  ', ' '));
        
    [list_raw, list_raw_class, list_ref, list_score_class, list_match_type, list_range_ref] = QualityAssessment_S3(data_org, total_features, c+c1, ranges, pos_metas, ranges_metas, total_features_1, total_features_metas);
    stop = timeit.default_timer();
    print('Time: ', np.around(stop - start, 3), 'sec'); print();
    
    print("Creating the evaluation report...");
    start = timeit.default_timer();
    write_evaluation_report(data_org, r, c, features_total, metas_features, pos_metas, ranges, var_type_final, var_type_final_2, 
                            var_type_metas, var_type_metas_2, features_state_metas, incompatibilities_metas, features_missing_values_metas, 
                            bad_features_metas, bad_features_ind_metas,fair_features_metas, fair_features_ind_metas, 
                            good_features_metas,good_features_ind_metas, a_total_metas, outliers_ind_metas, y_score_metas, 
                            outliers_pos_metas, ranges_metas, features_missing_values, features_state, outliers_ind, incompatibilities, a_total, path_f_e);
    stop = timeit.default_timer();
    print('Time: ', np.around(stop - start, 3), 'sec'); print();

    start = timeit.default_timer();
    print("Creating the standardization report...");
    write_standardization_report(total_features, list_raw, list_raw_class, list_ref, list_score_class, list_match_type, list_range_ref, path_f_s);
    stop = timeit.default_timer();
    print('Time: ', np.around(stop - start, 3), 'sec'); print();
    
    print("Creating the curated dataset...");
    start = timeit.default_timer();
    write_curated_dataset(data_org, wb, pos_metas, features_total, features_state, metas_features, imputation_method_id, 
                          var_type_final, outliers_pos, r, c, features_state_metas, outliers_pos_metas, var_type_metas, 
                          var_type_metas_2, y_total, incomp_pos, incomp_pos_metas, path_f_c);     
    stop = timeit.default_timer();
    print('Time: ', np.around(stop - start, 3), 'sec'); print();

#    plot_curator_results(data_org, cmatrix, outliers_ind, good_features_ind);

    print("Done!");


def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route('/upload', methods=['GET', 'POST'])
def index():
   if request.method == 'POST':
       if 'file' not in request.files:
           print('No file attached in request');
           return redirect(request.url);
       file = request.files['file'];
       if file.filename == '':
           print('No file selected');
           return redirect(request.url);
       if file and allowed_file(file.filename):
           filename = secure_filename(file.filename);
           file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename));
           imputation_method_id = request.form.get("imputation_method_id");
           outlier_detection_method_id = request.form.get("outlier_detection_method_id");
           
           if(imputation_method_id is None):
               imputation_method_id = 0;
        
           if(outlier_detection_method_id is None):
               outlier_detection_method_id = 5;       
        
           print('Method for imputation:', np.str(imputation_method_id));
           print('Method for outlier_detection:', np.str(outlier_detection_method_id));
           
           data_curator(os.path.join(app.config['UPLOAD_FOLDER'], filename), imputation_method_id, outlier_detection_method_id);
           os.remove(os.path.join(app.config['UPLOAD_FOLDER'], filename));
           return redirect(url_for('return_file'));
   return render_template('index.html');


@app.route('/ajax/index')
def ajax_index():
    time.sleep(5)
    return '<h1>Done!</h1>';


@app.route('/download')
def return_file():
    zipf = zipfile.ZipFile('Curator_results.zip','w', zipfile.ZIP_DEFLATED);
    for root,dirs, files in os.walk('results/'):
        for file in files:
            zipf.write('results/'+file);
            
    for root,dirs, files in os.walk('uploaded/'):
        for file in files:
            zipf.write('uploaded/'+file);
            
    zipf.close();
    return send_file('Curator_results.zip', as_attachment=True);


if __name__ == "__main__":
    app.run(host="195.130.121.195", port=14);
    