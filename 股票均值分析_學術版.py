#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import pandas as pd
import numpy as np
import requests
from io import StringIO
import time
import schedule
from datetime import datetime, timedelta
import os
import csv
import re
import calendar
import json
import logging
import sys
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import matplotlib.pyplot as plt
from matplotlib.font_manager import FontProperties
import matplotlib
matplotlib.use('Agg')  

class TWStockAnalyzer:
    def __init__(self, config_file='config.json'):
        """初始化分析器並讀取設定檔"""
        self.load_config(config_file)
        
        # 設置日誌
        log_dir = os.path.join(self.export_path, 'stock_logs')
        if not os.path.exists(log_dir):
            os.makedirs(log_dir)
        
        # 設置日誌格式
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(os.path.join(log_dir, f'stock_analyzer_{datetime.now().strftime("%Y%m%d")}.log')),
                logging.StreamHandler()
            ]
        )
        
        # 設置請求頭
        self.headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
        }
    
    def load_config(self, config_file):
        """載入設定檔"""
        try:
            with open(config_file, 'r', encoding='utf-8') as f:
                config = json.load(f)
            
            # 載入設定參數
            self.export_path = config.get('export_path', './output')
            self.export_filename = config.get('export_filename', 'tw_stock_ma_breakthrough_{date}.csv')
            self.run_time = config.get('run_time', '18:30')  # 保留但不再使用於排程
            self.holidays = [datetime.strptime(date, '%Y-%m-%d') for date in config.get('holidays', [])]
            
            # 確保匯出目錄存在
            if not os.path.exists(self.export_path):
                os.makedirs(self.export_path)
                
        except Exception as e:
            print(f"讀取設定檔時發生錯誤: {e}")
            print("使用預設設定，並建立 config.json")

            # 預設設定值
            self.export_path = './output'
            self.export_filename = 'tw_stock_ma_breakthrough_{date}.csv'
            self.run_time = '18:30'
            self.holidays = []

            if not os.path.exists(self.export_path):
                os.makedirs(self.export_path)
            
            # 寫入預設 config.json 檔案
            default_config = {
                "export_path": self.export_path,
                "export_filename": self.export_filename,
                "run_time": self.run_time,
                "holidays": []
            }
            try:
                with open(config_file, 'w', encoding='utf-8') as f:
                    json.dump(default_config, f, indent=4, ensure_ascii=False)
                print(f"已建立預設的設定檔：{config_file}")
            except Exception as write_error:
                print(f"寫入預設設定檔失敗: {write_error}")
        
    def is_trading_day(self, date):
        """檢查指定日期是否為交易日"""
        # 週末不交易
        if date.weekday() >= 5:  # 5是週六，6是週日
            return False
        
        # 檢查是否為假日
        if date in self.holidays:
            return False
            
        return True
    
    def get_previous_trading_day(self, date):
        """獲取指定日期的前一個交易日"""
        current = date - timedelta(days=1)
        while not self.is_trading_day(current):
            current = current - timedelta(days=1)
        return current
    
    def calculate_start_date(self, end_date, days_needed=60):
        """計算需要的起始日期，確保有足夠的交易日數據"""
        # 假設大約需要 days_needed 個交易日 (約三個月)
        current = end_date
        trading_days = 0
        calendar_days = 0
        
        while trading_days < days_needed:
            current = current - timedelta(days=1)
            calendar_days += 1
            if self.is_trading_day(current):
                trading_days += 1
        
        # 再多加一些天以確保足夠
        buffer_days = 10
        return current - timedelta(days=buffer_days)
        
    def fetch_twse_data(self, date):
        """獲取台灣證券交易所(TWSE)上市公司的每日收盤價資料"""
        year = date.year
        month = date.month
        day = date.day
        
        # 將日期轉換為台灣證交所要求的格式
        date_str = f"{year}{month:02d}{day:02d}"
        
        # 台灣證券交易所API的URL - 使用較穩定的每日收盤行情
        url = f"https://www.twse.com.tw/exchangeReport/MI_INDEX?response=csv&date={date_str}&type=ALLBUT0999"
        
        try:
            # 使用stream=True來處理不同的編碼
            response = requests.get(url, headers=self.headers, stream=True)
            response.raise_for_status()  # 檢查請求是否成功
            
            # 嘗試不同的編碼
            encodings = ['utf-8', 'big5', 'big5hkscs', 'cp950']
            content = None
            
            for encoding in encodings:
                try:
                    content = response.content.decode(encoding)
                    break
                except UnicodeDecodeError:
                    continue
            
            if content is None:
                logging.error(f"無法解碼 {date_str} 的TWSE數據")
                return pd.DataFrame()
            
            # 尋找股票資料的部分
            lines = content.split('\n')
            start_idx = -1
            for i, line in enumerate(lines):
                if '"證券代號"' in line and '"證券名稱"' in line and '"收盤價"' in line:
                    start_idx = i
                    break
            
            if start_idx == -1:
                logging.error(f"無法找到 {date_str} 的TWSE股票數據表格")
                return pd.DataFrame()
            
            # 提取有效的資料行
            data_lines = []
            data_lines.append(lines[start_idx])  # 加入標題行
            
            for line in lines[start_idx+1:]:
                # 檢查行是否包含股票數據 (通常以數字開頭)
                if line.strip() and re.match(r'^"[0-9]{4,}"', line.strip()):
                    data_lines.append(line)
                elif line.strip() and '==================================' in line:
                    break  # 發現分隔線，表示數據結束
            
            # 如果沒有數據，返回空DataFrame
            if len(data_lines) <= 1:
                logging.warning(f"TWSE在 {date_str} 未找到股票數據")
                return pd.DataFrame()
            
            # 手動解析CSV數據
            csv_data = []
            header = None
            
            for i, line in enumerate(data_lines):
                # 使用CSV模塊來正確處理CSV格式
                reader = csv.reader(StringIO(line), delimiter=',', quotechar='"')
                row = next(reader)
                
                if i == 0:  # 這是標題行
                    header = [col.strip('"').strip() for col in row]
                    # 查找必要列的索引
                    try:
                        code_idx = header.index('證券代號')
                        name_idx = header.index('證券名稱')
                        close_idx = header.index('收盤價')
                    except ValueError:
                        logging.error(f"TWSE數據缺少必要的列: {header}")
                        return pd.DataFrame()
                else:
                    if len(row) >= max(code_idx, name_idx, close_idx) + 1:
                        csv_data.append({
                            'stock_id': row[code_idx].strip('"'),
                            'stock_name': row[name_idx].strip('"'),
                            'close': row[close_idx].strip('"').replace(',', '')
                        })
            
            # 創建DataFrame
            df = pd.DataFrame(csv_data)
            
            # 確保收盤價是數值型
            df['close'] = pd.to_numeric(df['close'], errors='coerce')
            
            # 過濾掉非數值的行及特殊股票
            df = df.dropna(subset=['close'])
            df = df[df['stock_id'].str.isdigit()]  # 只保留數字股票代碼
            
            # 設置日期列
            df['date'] = date
            
            return df
                
        except Exception as e:
            logging.error(f"獲取TWSE數據時發生錯誤: {e}")
            return pd.DataFrame()
    
    def fetch_tpex_data(self, date):
        """獲取證券櫃檯買賣中心(TPEx)上櫃公司的每日收盤價資料"""
        year = date.year - 1911  # 轉換為民國年
        month = date.month
        day = date.day
        
        # 證券櫃檯買賣中心API的URL - 使用較穩定的CSV下載連結
        url = f"https://www.tpex.org.tw/web/stock/aftertrading/daily_close_quotes/stk_quote_download.php?l=zh-tw&d={year}/{month:02d}/{day:02d}&s=0,asc,0"
        
        try:
            # 使用stream=True來處理不同的編碼
            response = requests.get(url, headers=self.headers, stream=True)
            response.raise_for_status()  # 檢查請求是否成功
            
            # 嘗試不同的編碼
            encodings = ['utf-8', 'big5', 'big5hkscs', 'cp950']
            content = None
            
            for encoding in encodings:
                try:
                    content = response.content.decode(encoding)
                    break
                except UnicodeDecodeError:
                    continue
            
            if content is None:
                logging.error(f"無法解碼 {date} 的TPEx數據")
                return pd.DataFrame()
            
            # 檢查是否有數據
            if "查無資料" in content or "請重新查詢" in content or len(content.strip()) == 0:
                logging.warning(f"TPEx在 {date} 沒有數據")
                return pd.DataFrame()
            
            # 手動解析CSV內容
            lines = content.split('\n')
            header_found = False
            data_start = False
            data_lines = []
            
            for i, line in enumerate(lines):
                # 跳過空行
                if not line.strip():
                    continue
                    
                # 找到標題行
                if not header_found and ("代號" in line and "名稱" in line and "收盤" in line):
                    header_found = True
                    data_lines.append(line)
                    data_start = True
                    continue
                
                # 收集數據行
                if data_start:
                    # 檢查是否到達數據結束
                    if "總計" in line or "加權指數" in line:
                        break
                    # 確保這是有效的數據行 (通常以數字開頭)
                    if re.match(r'^[0-9]{4,}', line.strip().split(',')[0].strip()):
                        data_lines.append(line)
            
            # 如果沒有找到有效數據
            if not header_found or len(data_lines) <= 1:
                logging.warning(f"TPEx在 {date} 無法找到有效的股票數據")
                return pd.DataFrame()
            
            # 手動解析CSV數據
            csv_data = []
            header = data_lines[0].split(',')
            
            # 找尋必要列的索引
            try:
                code_idx = 0  # 第一列通常是代號
                name_idx = 1  # 第二列通常是名稱
                close_idx = 2  # 第三列通常是收盤價
            except ValueError:
                logging.error(f"TPEx數據缺少必要的列: {header}")
                return pd.DataFrame()
            
            # 解析數據行
            for i in range(1, len(data_lines)):
                fields = data_lines[i].split(',')
                if len(fields) >= max(code_idx, name_idx, close_idx) + 1:
                    stock_id = fields[code_idx].strip()
                    # 確保這是有效的股票代碼
                    if stock_id.isdigit():
                        csv_data.append({
                            'stock_id': stock_id,
                            'stock_name': fields[name_idx].strip(),
                            'close': fields[close_idx].strip().replace(',', '')
                        })
            
            # 創建DataFrame
            df = pd.DataFrame(csv_data)
            
            # 確保收盤價是數值型
            df['close'] = pd.to_numeric(df['close'], errors='coerce')
            
            # 過濾掉非數值的行
            df = df.dropna(subset=['close'])
            
            # 設置日期列
            df['date'] = date
            
            return df
                
        except Exception as e:
            logging.error(f"獲取TPEx數據時發生錯誤: {e}")
            return pd.DataFrame()
    
    def fetch_data_for_date_range(self, start_date, end_date, max_retry=3):
        """獲取指定日期範圍內的所有股票數據"""
        all_data = []
        
        # 確保日期格式正確
        if isinstance(start_date, str):
            start_date = datetime.strptime(start_date, '%Y-%m-%d')
        if isinstance(end_date, str):
            end_date = datetime.strptime(end_date, '%Y-%m-%d')
        
        # 計算總交易日
        total_trading_days = 0
        temp_date = start_date
        while temp_date <= end_date:
            if self.is_trading_day(temp_date):
                total_trading_days += 1
            temp_date += timedelta(days=1)
        
        logging.info(f"從 {start_date.strftime('%Y-%m-%d')} 到 {end_date.strftime('%Y-%m-%d')} 預計有 {total_trading_days} 個交易日")
        
        # 獲取每個交易日的數據
        current_date = start_date
        processed_days = 0
        retry_dates = []  # 存儲需要重試的日期
        
        while current_date <= end_date:
            # 如果不是交易日，跳過
            if not self.is_trading_day(current_date):
                current_date += timedelta(days=1)
                continue
            
            processed_days += 1
            logging.info(f"[{processed_days}/{total_trading_days}] 獲取 {current_date.strftime('%Y-%m-%d')} 的數據...")
            
            # 獲取TWSE數據
            twse_data = self.fetch_twse_data(current_date)
            if not twse_data.empty:
                logging.info(f"  - 成功獲取TWSE數據: {len(twse_data)}筆")
                twse_data['market'] = 'TWSE'
                all_data.append(twse_data)
            else:
                retry_dates.append((current_date, 'TWSE'))
            
            # 獲取TPEx數據
            tpex_data = self.fetch_tpex_data(current_date)
            if not tpex_data.empty:
                logging.info(f"  - 成功獲取TPEx數據: {len(tpex_data)}筆")
                tpex_data['market'] = 'TPEx'
                all_data.append(tpex_data)
            else:
                retry_dates.append((current_date, 'TPEx'))
            
            # 加入延遲以避免過多請求
            time.sleep(2)
            
            current_date += timedelta(days=1)
        
        # 重試失敗的日期
        if retry_dates and max_retry > 0:
            logging.info(f"\n開始重試 {len(retry_dates)} 個失敗的請求...")
            for retry_date, market in retry_dates:
                logging.info(f"重試獲取 {retry_date.strftime('%Y-%m-%d')} 的 {market} 數據...")
                
                if market == 'TWSE':
                    data = self.fetch_twse_data(retry_date)
                    if not data.empty:
                        logging.info(f"  - 重試成功獲取TWSE數據: {len(data)}筆")
                        data['market'] = 'TWSE'
                        all_data.append(data)
                else:
                    data = self.fetch_tpex_data(retry_date)
                    if not data.empty:
                        logging.info(f"  - 重試成功獲取TPEx數據: {len(data)}筆")
                        data['market'] = 'TPEx'
                        all_data.append(data)
                
                time.sleep(2)
        
        # 合併所有數據
        if all_data:
            combined_data = pd.concat(all_data, ignore_index=True)
            logging.info(f"總共獲取了 {len(combined_data)} 筆股票數據")
            return combined_data
        else:
            logging.warning("未獲取到有效數據")
            return pd.DataFrame()
    
    def calculate_moving_averages(self, data, windows=[5, 10, 20]):
        """計算指定窗口的移動平均線"""
        # 按股票ID和日期排序
        data = data.sort_values(['stock_id', 'date'])
        
        # 創建結果DataFrame的副本
        result = data.copy()
        
        # 初始化MA列
        for window in windows:
            result[f'MA{window}'] = np.nan
        
        # 檢查數據量
        stock_dates = data.groupby('stock_id')['date'].nunique()
        logging.info(f"數據中包含的股票數量: {len(stock_dates)}")
        logging.info(f"每支股票的平均交易日數: {stock_dates.mean():.2f}")
        logging.info(f"最小交易日數: {stock_dates.min()}, 最大交易日數: {stock_dates.max()}")
        
        # 打印具體的日期範圍
        min_date = data['date'].min()
        max_date = data['date'].max()
        logging.info(f"數據日期範圍: {min_date} 到 {max_date}")
        
        # 計算每支股票的均線
        stocks_with_insufficient_data = []
        stock_groups = data.groupby('stock_id')
        total_stocks = len(data['stock_id'].unique())
        processed = 0
        
        for stock_id, stock_data in stock_groups:
            # 對每個窗口計算移動平均線
            stock_df = stock_data.sort_values('date')
            days_count = len(stock_df)
            
            for window in windows:
                ma_col = f'MA{window}'
                if days_count < window:
                    # 記錄數據不足的股票
                    if stock_id not in stocks_with_insufficient_data:
                        stocks_with_insufficient_data.append(stock_id)
                    # 使用完整的window值，但確保min_periods設為window
                    # 這樣在數據不足時會產生NaN而不是部分計算的值
                    result.loc[stock_df.index, ma_col] = np.nan
                else:
                    # 確保計算正確，設置min_periods=window
                    ma_values = stock_df['close'].rolling(window=window, min_periods=window).mean()
                    result.loc[stock_df.index, ma_col] = ma_values.values
            
            processed += 1
            if processed % 100 == 0 or processed == total_stocks:
                logging.info(f"已處理 {processed}/{total_stocks} 支股票的均線計算")
        
        if stocks_with_insufficient_data:
            logging.warning(f"警告: {len(stocks_with_insufficient_data)} 支股票的數據少於 {max(windows)} 個交易日")
            logging.warning(f"這些股票的均線計算可能不准確或為NaN")
        
        logging.info("移動平均線計算完成")
        return result
    
    def filter_stocks(self, data, date1, date2):
        """篩選符合條件的股票"""
        if data.empty:
            logging.warning("沒有數據可供篩選")
            return pd.DataFrame()
            
        # 確保日期格式正確
        if isinstance(date1, str):
            date1 = datetime.strptime(date1, '%Y-%m-%d')
        if isinstance(date2, str):
            date2 = datetime.strptime(date2, '%Y-%m-%d')
        
        logging.info(f"開始篩選符合條件的股票...")
        logging.info(f"可用日期範圍: {data['date'].min()} 到 {data['date'].max()}")
        
        # 第一個日期：收盤價低於均線
        date1_data = data[data['date'] == date1].copy()
        if date1_data.empty:
            logging.warning(f"警告：找不到 {date1.strftime('%Y-%m-%d')} 的數據")
            return pd.DataFrame()
            
        date1_filtered = date1_data[(date1_data['close'] < date1_data['MA5']) & 
                                   (date1_data['close'] < date1_data['MA10']) &
                                   (date1_data['close'] < date1_data['MA20'])]
        logging.info(f"找到 {len(date1_filtered)} 支股票在 {date1.strftime('%Y-%m-%d')} 收盤價低於均線")
        
        # 第二個日期：收盤價高於均線
        date2_data = data[data['date'] == date2].copy()
        if date2_data.empty:
            logging.warning(f"警告：找不到 {date2.strftime('%Y-%m-%d')} 的數據")
            return pd.DataFrame()
            
        date2_filtered = date2_data[(date2_data['close'] > date2_data['MA5']) & 
                                   (date2_data['close'] > date2_data['MA10']) &
                                   (date2_data['close'] > date2_data['MA20'])]
        logging.info(f"找到 {len(date2_filtered)} 支股票在 {date2.strftime('%Y-%m-%d')} 收盤價高於均線")
        
        # 找出同時符合兩個條件的股票
        common_stocks = set(date1_filtered['stock_id']).intersection(set(date2_filtered['stock_id']))
        logging.info(f"找到 {len(common_stocks)} 支股票符合均線突破條件")
        
        # 如果沒有找到符合條件的股票
        if not common_stocks:
            return pd.DataFrame()
        
        # 獲取這些股票的詳細信息
        result = []
        for stock_id in common_stocks:
            try:
                stock_info_date1 = date1_filtered[date1_filtered['stock_id'] == stock_id].iloc[0]
                stock_info_date2 = date2_filtered[date2_filtered['stock_id'] == stock_id].iloc[0]
                
                result.append({
                    'stock_id': stock_id,
                    'stock_name': stock_info_date1['stock_name'],
                    'market': stock_info_date1['market'],
                    f'close_{date1.strftime("%Y%m%d")}': stock_info_date1['close'],
                    f'MA5_{date1.strftime("%Y%m%d")}': stock_info_date1['MA5'],
                    f'MA10_{date1.strftime("%Y%m%d")}': stock_info_date1['MA10'],
                    f'MA20_{date1.strftime("%Y%m%d")}': stock_info_date1['MA20'],
                    f'close_{date2.strftime("%Y%m%d")}': stock_info_date2['close'],
                    f'MA5_{date2.strftime("%Y%m%d")}': stock_info_date2['MA5'],
                    f'MA10_{date2.strftime("%Y%m%d")}': stock_info_date2['MA10'],
                    f'MA20_{date2.strftime("%Y%m%d")}': stock_info_date2['MA20']
                })
            except IndexError:
                logging.warning(f"警告：處理股票 {stock_id} 時出現索引錯誤，可能是數據不完整")
                continue
        
        return pd.DataFrame(result)
    
    def run_analysis(self, output_file=None):
        """執行完整的分析流程，自動使用當天和前一個交易日"""
        # 獲取當前日期和前一個交易日
        today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
        
        # 檢查今天是否為交易日，如果不是，找最近的交易日
        if not self.is_trading_day(today):
            latest_trading_day = today
            while not self.is_trading_day(latest_trading_day):
                latest_trading_day = latest_trading_day - timedelta(days=1)
            logging.info(f"今天 {today.strftime('%Y-%m-%d')} 不是交易日，使用最近的交易日 {latest_trading_day.strftime('%Y-%m-%d')}")
            today = latest_trading_day
        
        # 獲取前一個交易日
        previous_trading_day = self.get_previous_trading_day(today)
        
        logging.info(f"分析日期: 今天 {today.strftime('%Y-%m-%d')} 和前一交易日 {previous_trading_day.strftime('%Y-%m-%d')}")
        
        # 自動計算起始日期，確保有足夠的交易日計算MA20
        start_date = self.calculate_start_date(today, days_needed=60)  # 約三個月60個交易日
        logging.info(f"自動計算的起始日期: {start_date.strftime('%Y-%m-%d')}")
        
        # 產生輸出檔案名稱
        if output_file is None:
            date_str = today.strftime('%Y%m%d')
            output_file = os.path.join(
                self.export_path, 
                self.export_filename.format(date=date_str)
            )
        
        # 獲取歷史數據
        logging.info("開始獲取歷史數據...")
        data = self.fetch_data_for_date_range(start_date, today)
        
        if data.empty:
            logging.error("未獲取到有效數據，分析終止")
            return None
        
        # 儲存原始數據以備後用
        raw_data_file = os.path.join(self.export_path, f'raw_stock_data_{today.strftime("%Y%m%d")}.csv')
        data.to_csv(raw_data_file, index=False, encoding='utf_8_sig')
        logging.info(f"原始數據已保存至 {raw_data_file}")
        
        # 計算移動平均線
        logging.info("計算移動平均線...")
        data_with_ma = self.calculate_moving_averages(data)
        
        # 儲存含MA的數據以備後用
        ma_data_file = os.path.join(self.export_path, f'stock_data_with_ma_{today.strftime("%Y%m%d")}.csv')
        data_with_ma.to_csv(ma_data_file, index=False, encoding='utf_8_sig')
        logging.info(f"含均線的數據已保存至 {ma_data_file}")
        
        # 篩選符合條件的股票
        logging.info("篩選符合條件的股票...")
        filtered_stocks = self.filter_stocks(data_with_ma, previous_trading_day, today)
        
        # 保存結果
        pdf_file = None
        if not filtered_stocks.empty:
            filtered_stocks.to_csv(output_file, index=False, encoding='utf_8_sig')
            logging.info(f"分析結果已保存至 {output_file}")
            logging.info(f"找到 {len(filtered_stocks)} 支符合條件的股票")
            # 列出找到的股票
            for idx, row in filtered_stocks.iterrows():
                logging.info(f"{row['stock_id']} - {row['stock_name']} ({row['market']})")

            # 生成並保存圖表
            pdf_file = self.generate_chart(filtered_stocks)

            return filtered_stocks, pdf_file
        else:
            logging.warning("未找到符合條件的股票")
            return None, None
        
    def generate_chart(self, filtered_stocks, output_pdf=None):
        """生成均線突破股票的圖表，並保存為PDF"""
        if filtered_stocks is None or filtered_stocks.empty:
            logging.warning("沒有數據可用於生成圖表")
            return None

        # 如果沒有指定輸出文件，則使用默認命名
        if output_pdf is None:
            date_str = datetime.now().strftime('%Y%m%d')
            output_pdf = os.path.join(
                self.export_path, 
                f"tw_stock_ma_breakthrough_chart_{date_str}.pdf"
            )

        # 設置中文字體
        try:
            # 嘗試使用系統中文字體
            font_path = None
            if os.path.exists("/usr/share/fonts/truetype/arphic/uming.ttc"):  # Linux
                font_path = "/usr/share/fonts/truetype/arphic/uming.ttc"
            elif os.path.exists("/System/Library/Fonts/PingFang.ttc"):  # macOS
                font_path = "/System/Library/Fonts/PingFang.ttc"
            elif os.path.exists("C:\\Windows\\Fonts\\msjh.ttc"):  # Windows
                font_path = "C:\\Windows\\Fonts\\msjh.ttc"

            if font_path:
                plt.rcParams['font.family'] = 'sans-serif'
                plt.rcParams['font.sans-serif'] = ['Arial Unicode MS', 'SimHei']
                font = FontProperties(fname=font_path)
            else:
                # 如果沒有找到合適的中文字體，使用默認設置
                plt.rcParams['font.sans-serif'] = ['Arial Unicode MS', 'SimHei', 'DejaVu Sans']
                font = None
        except Exception as e:
            logging.warning(f"設置中文字體時出錯：{e}，將使用默認字體")
            font = None

        # 創建圖表
        plt.figure(figsize=(12, 8))

        # 獲取日期字符串（用於列標題）
        date_cols = [col for col in filtered_stocks.columns if 'close_' in col]
        date_strs = [col.replace('close_', '') for col in date_cols]

        # 對股票按照價格變化百分比排序
        filtered_stocks['price_change_pct'] = ((filtered_stocks[f'close_{date_strs[1]}'] - 
                                               filtered_stocks[f'close_{date_strs[0]}']) / 
                                              filtered_stocks[f'close_{date_strs[0]}'] * 100)

        sorted_stocks = filtered_stocks.sort_values('price_change_pct', ascending=False)

        # 繪製價格變化百分比圖表
        bars = plt.bar(range(len(sorted_stocks)), 
                       sorted_stocks['price_change_pct'], 
                       color='royalblue')

        # 添加股票代碼和名稱標籤（x 軸）
        stock_labels = [f"{row['stock_id']}\n{row['stock_name']}" for _, row in sorted_stocks.iterrows()]
        plt.xticks(range(len(sorted_stocks)), stock_labels, rotation=90, fontproperties=font)

        # 添加數值標籤（柱上百分比）
        for i, bar in enumerate(bars):
            height = bar.get_height()
            plt.text(bar.get_x() + bar.get_width()/2., 
                     height + 0.3,
                     f"{height:.2f}%", 
                     ha='center', va='bottom', rotation=0,
                     fontproperties=font)  # <- 加入這行！

        # 設定標題與 y 軸標籤
        plt.title(f"均線突破股票價格變化百分比 ({date_strs[0]}→{date_strs[1]})", fontproperties=font)
        plt.ylabel("價格變化百分比 (%)", fontproperties=font)
        plt.grid(axis='y', linestyle='--', alpha=0.7)

        # 調整版面
        plt.tight_layout()

        # 保存為PDF
        try:
            plt.savefig(output_pdf)
            logging.info(f"圖表已保存為PDF：{output_pdf}")
            plt.close()
            return output_pdf
        except Exception as e:
            logging.error(f"保存圖表PDF時出錯：{e}")
            plt.close()
            return None
    
    def run_once(self):
        """立即執行一次分析"""
        logging.info("立即執行一次分析")
        try:
            result, pdf_file = self.run_analysis()
            return result, pdf_file
        except Exception as e:
            logging.error(f"執行分析時發生錯誤: {e}")
            return None, None
        
def sendemail(to, sub, context, attachments=None):
    config = readconfig()
    
    if not config:
        print("無法讀取郵件配置")
        return False
    
    send_email = config["send_email"]
    send_password = config["send_password"]
    smtp = config.get("smtp", "smtp.gmail.com")
    smtp_port = config.get("smtp_port", 587)
    
    # 創建郵件物件
    email = MIMEMultipart()
    email["From"] = send_email
    email["To"] = to
    email["Subject"] = sub
    
    # 添加郵件內容
    email.attach(MIMEText(context, "plain"))
    
    # 如果有附件，則添加
    if attachments:
        if isinstance(attachments, str):
            attachments = [attachments]  # 轉換單一附件為列表
            
        for attachment_path in attachments:
            if attachment_path and os.path.exists(attachment_path):
                with open(attachment_path, "rb") as file:
                    filename = os.path.basename(attachment_path)
                    filepart = MIMEApplication(file.read(), Name=filename)
                    filepart["Content-Disposition"] = f'attachment; filename="{filename}"'
                    email.attach(filepart)
    
    try:
        # 連接到SMTP伺服器
        server = smtplib.SMTP(smtp, smtp_port)
        server.starttls()  # 啟用TLS加密
        server.login(send_email, send_password)
        
        # 發送郵件
        server.send_message(email)
        server.quit()
        
        print(f"郵件成功寄送給 {to}")
        return True
    except Exception as e:
        print(f"寄送郵件時出錯: {e}")
        return False
    
def file_exist(expath, filename):
    # 替換 {date} 占位符為當天日期
    today = datetime.now().strftime('%Y%m%d')  # 或者你想要的日期格式
    filename = filename.replace('{date}', today)
    
    complete_path = os.path.join(expath, filename)
    return os.path.exists(complete_path), complete_path, filename

def work():
    if __name__ == "__main__":
        try:
            # 創建股票分析器實例
            analyzer = TWStockAnalyzer()

            # 檢查命令行參數
            if len(sys.argv) > 1 and sys.argv[1] == "--schedule":
                # 這裡可以添加排程功能
                print("排程功能尚未實現，請使用系統排程工具如cron或Windows排程器")
                sys.exit(0)
            else:
                # 立即執行一次分析
                print("開始執行股票均線突破分析...")
                result, pdf_file = analyzer.run_once()
                if result is not None:
                    print(f"分析完成，找到 {len(result)} 支符合條件的股票")
                else:
                    print("分析完成，未找到符合條件的股票或執行過程中出現錯誤")
                config = readconfig()
                if not config:
                    return False
    
                # 從配置中獲取設置
                expath = config["export_path"]
                filename = config.get('export_filename', 'tw_stock_ma_breakthrough_{date}.csv')
                to = config["to"]

                sub = config["sub"]
                context = config["context"]

                nofile_sub = config["nofile_sub"]
                nofile_context = config["nofile_context"]
        
                fileexist, completepath, formatted_filename = file_exist(expath, filename)

                # 檢查檔案是否存在
                if fileexist:
                    print(f"找到檔案: {completepath}")
                    # 使用有檔案的郵件內容，並添加附件（CSV和PDF）
                    attachments = [completepath]
                    if pdf_file and os.path.exists(pdf_file):
                        attachments.append(pdf_file)
                        print(f"添加圖表PDF附件: {pdf_file}")
                    sendemail(to, sub, context, attachments)
                else:
                    print(f"找不到檔案: {expath}/{formatted_filename}")
                    # 使用無檔案的郵件內容，檢查是否有PDF可附加
                    if pdf_file and os.path.exists(pdf_file):
                        print(f"添加圖表PDF附件: {pdf_file}")
                        sendemail(to, nofile_sub, nofile_context, pdf_file)
                    else:
                        sendemail(to, nofile_sub, nofile_context)
                
        except Exception as e:
            print(f"程式執行時發生未預期的錯誤: {e}")
            logging.error(f"程式執行時發生未預期的錯誤: {e}", exc_info=True)
            sys.exit(1)
            
def readconfig():
    try:
        with open('config.json', 'r', encoding='utf-8') as file:
            config = json.load(file)
        return config
    except Exception as e:
        print(f"讀取配置檔案時出錯: {e}")
        return None

def setschedule():
    # 讀取配置檔案
    config = readconfig()
    
    if config and "run_time" in config:
        set_time = config["run_time"]
        print(f"從配置檔案讀取的執行時間: {set_time}")
        
        # 清除之前的所有排程
        schedule.clear()
        
        # 設定新的排程
        schedule.every().day.at(set_time).do(work)
        print(f"已設置每天 {set_time} 執行任務")
    else:
        print("無法從配置檔案讀取執行時間")

if __name__ == "__main__":
    print("開始監控排程...")
    setschedule()  # 初始設置排程
    
    # 添加無限循環來運行排程任務
    while True:
        schedule.run_pending()  # 運行待執行的任務
        time.sleep(1)  # 休眠1秒，避免CPU使用率過高


# In[ ]:




