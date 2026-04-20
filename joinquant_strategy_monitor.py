"""
策略环境监测系统 - 聚宽平台版本
整合指标计算和策略评分功能，支持Excel输出
"""
import numpy as np
import pandas as pd
import datetime as dt  # 使用别名避免冲突
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# 聚宽数据API导入
from jqdata import *

# 评分用阈值与加减分：不在代码里写死数值，避免用不准确的常数代替真实数据。
# 需要打分时在此字典中填写，例如：
# SCORE_WEIGHTS = {
#     'subjective': {
#         'market_amount_high': (10000, 12),   # (阈值亿元, 加分)
#         'market_amount_low': (6000, -12),
#         'pe_pct_cheap': (30, 18), 'pe_pct_expensive': (70, -18),
#         'gv_ratio_high': (2, 10), 'gv_ratio_low': (-2, -10),
#         'amount_change_high': (10, 8), 'amount_change_low': (-10, -8),
#     },
#     'quant': {
#         'market_amount_high': (10000, 10), 'market_amount_low': (6000, -10),
#         'crowding_high': (20, -15), 'crowding_low': (10, 15),
#     },
# }
# 各策略评分函数在 SCORE_WEIGHTS 为空时一律保持 50 分并列明原因，避免用硬编码常数冒充数据。
SCORE_WEIGHTS = {}

# ============================================
# 数据获取模块 - 使用聚宽API
# ============================================

class DataFetcher:
    """数据获取类 - 聚宽版本"""
    
    def __init__(self):
        pass
    
    def get_index_daily(self, symbol, period=120, end_date=None):
        """获取指数日线数据"""
        if end_date is None:
            end_date = dt.datetime.now().strftime('%Y-%m-%d')
        
        try:
            df = get_price(symbol, count=period, end_date=end_date, frequency='daily', fields=['open', 'close', 'high', 'low', 'volume', 'money'])
            if df is not None and not df.empty:
                df = df.reset_index()
                df.rename(columns={'index': '日期', 'open': '开盘', 'close': '收盘', 
                                  'high': '最高', 'low': '最低', 'volume': '成交量', 
                                  'money': '成交额'}, inplace=True)
                df['日期'] = pd.to_datetime(df['日期'])
            return df
        except Exception as e:
            print(f"获取指数数据失败 {symbol}: {e}")
            return None
    
    def get_etf_daily(self, symbol, period=252, end_date=None):
        """获取ETF日线数据"""
        if end_date is None:
            end_date = dt.datetime.now().strftime('%Y-%m-%d')
        
        try:
            df = get_price(symbol, count=period, end_date=end_date, frequency='daily', fields=['open', 'close', 'high', 'low', 'volume', 'money'])
            if df is not None and not df.empty:
                df = df.reset_index()
                df.rename(columns={'index': '日期', 'open': '开盘', 'close': '收盘',
                                  'high': '最高', 'low': '最低', 'volume': '成交量',
                                  'money': '成交额'}, inplace=True)
                df['日期'] = pd.to_datetime(df['日期'])
            return df
        except Exception as e:
            print(f"获取ETF数据失败 {symbol}: {e}")
            return None
    
    def get_etf_nav(self, symbol, period=252, end_date=None):
        """获取ETF单位净值数据 - 使用聚宽get_extras接口"""
        if end_date is None:
            end_date = dt.datetime.now().strftime('%Y-%m-%d')
        
        try:
            start_date = (dt.datetime.strptime(end_date, '%Y-%m-%d') - dt.timedelta(days=period)).strftime('%Y-%m-%d')
            # 使用get_extras获取ETF净值
            df = get_extras('unit_net_value', [symbol], start_date=start_date, end_date=end_date)
            if df is not None and not df.empty:
                df = df.reset_index()
                df.rename(columns={'index': '日期', symbol: '单位净值'}, inplace=True)
                df['日期'] = pd.to_datetime(df['日期'])
                # 去除NaN值
                df = df.dropna(subset=['单位净值'])
            return df
        except Exception as e:
            print(f"获取ETF净值数据失败 {symbol}: {e}")
            return None
    
    def get_futures_main_contract(self, underlying_symbol, period=252):
        """获取期货主力合约数据"""
        try:
            dominant = get_dominant_future(underlying_symbol)
            if dominant:
                ed = dt.datetime.now().strftime('%Y-%m-%d')
                df = get_price(dominant, count=period, end_date=ed, frequency='daily', fields=['open', 'close', 'high', 'low', 'volume', 'money'])
                if df is not None and not df.empty:
                    df = df.reset_index()
                    df.rename(columns={'index': '日期', 'open': '开盘价', 'close': '收盘价',
                                      'high': '最高价', 'low': '最低价', 'volume': '成交量',
                                      'money': '成交额'}, inplace=True)
                    df['最新价'] = df['收盘价']
                return df
            return None
        except Exception as e:
            print(f"获取期货主力合约失败 {underlying_symbol}: {e}")
            return None
    
    def get_futures_info_data(self, underlying_symbol):
        """获取期货合约信息和持仓量"""
        try:
            dominant = get_dominant_future(underlying_symbol)
            if dominant:
                df = get_extras('futures_positions', [dominant], 
                               start_date=(dt.datetime.now() - dt.timedelta(days=252)).strftime('%Y-%m-%d'),
                               end_date=dt.datetime.now().strftime('%Y-%m-%d'))
                return df
            return None
        except Exception as e:
            print(f"获取期货持仓量失败 {underlying_symbol}: {e}")
            return None


# ============================================
# 指标计算模块
# ============================================

class IndicatorCalculator:
    """指标计算类"""
    
    def __init__(self):
        self.fetcher = DataFetcher()
    
    def calculate_percentile_with_days(self, series, window=252, min_days=20):
        """计算历史分位数，并返回实际使用的天数
        :param series: 数据序列
        :param window: 最大窗口期（默认252天）
        :param min_days: 最少需要的天数（默认20天）
        :return: (分位数, 实际使用天数) 或 (None, 0) 如果数据不足
        """
        if series is None or len(series) == 0:
            return None, 0
        
        # 去除NaN值
        series_clean = series.dropna()
        
        # 取最近window天的数据
        hist = series_clean.tail(window)
        actual_days = len(hist)
        
        # 如果数据不足min_days天，不计算分位数
        if actual_days < min_days:
            return None, actual_days
        
        current = series.iloc[-1]
        percentile = (hist < current).mean() * 100
        return round(percentile, 2), actual_days
    
    # ==================== 主观多头指标 ====================
    
    def get_sentiment_indicators(self):
        """获取市场情绪指标"""
        indicators = {}
        
        # 0. 主要指数当日涨跌幅
        index_returns_map = {
            '沪深300': '000300.XSHG',
            '中证A500': '000510.XSHG',
            '中证500': '000905.XSHG',
            '中证1000': '000852.XSHG',
            '国证2000': '399303.XSHE',  # 小盘股指数
        }
        
        for name, code in index_returns_map.items():
            try:
                df = self.fetcher.get_index_daily(code, period=2)
                if df is not None and not df.empty and len(df) >= 2:
                    df = df.sort_values('日期')
                    daily_return = (df['收盘'].iloc[-1] / df['收盘'].iloc[-2] - 1) * 100
                    indicators[f'{name}_daily_return'] = round(daily_return, 2)
            except Exception as e:
                print(f"获取{name}当日涨跌幅失败: {e}")
        
        # 1. 交投活跃度 - 全市场成交额（以上证指数代理）
        try:
            df_index = self.fetcher.get_index_daily("000001.XSHG", period=40)
            if df_index is not None and not df_index.empty:
                df_index = df_index.sort_values('日期')
                avg_amount_20 = df_index['成交额'].tail(20).mean() / 1e8
                indicators['market_amount_20d_avg'] = round(avg_amount_20, 2)
                
                # 成交额变化率
                amount_change = (df_index['成交额'].iloc[-1] / df_index['成交额'].iloc[-20] - 1) * 100
                indicators['market_amount_change_20d'] = round(amount_change, 2)
        except Exception as e:
            print(f"计算市场成交额失败: {e}")
        
        return indicators
    
    def get_valuation_indicators(self):
        """获取估值指标 - 使用get_fundamentals获取指数成分股估值数据"""
        indicators = {}
        
        # 指数代码映射
        index_codes = {
            '沪深300': '000300.XSHG',
            '中证500': '000905.XSHG',
            '中证1000': '000852.XSHG',
            '国证2000': '399303.XSHE',
        }
        
        for name, code in index_codes.items():
            try:
                # 获取当前日期
                end_date = dt.datetime.now().strftime('%Y-%m-%d')
                
                # 获取指数成分股
                stocks = get_index_stocks(code)
                
                if stocks and len(stocks) > 0:
                    # 使用get_fundamentals获取成分股估值数据
                    q = query(
                        valuation.code,
                        valuation.pe_ratio,
                        valuation.pb_ratio,
                        valuation.market_cap
                    ).filter(
                        valuation.pe_ratio != None,
                        valuation.pb_ratio != None,
                        valuation.market_cap != None,
                        valuation.code.in_(stocks)
                    )
                    
                    df_val = get_fundamentals(q, end_date)
                    
                    if df_val is not None and not df_val.empty:
                        # 过滤掉无效数据
                        df_val = df_val.dropna(subset=['pe_ratio', 'pb_ratio', 'market_cap'])
                        df_val = df_val[df_val['pe_ratio'] > 0]  # 过滤亏损股
                        
                        if len(df_val) > 0:
                            # 计算指数PE = 总市值 / 总净利润
                            # 其中：总净利润 = Σ(个股市值 / 个股PE)
                            total_market_cap = df_val['market_cap'].sum()
                            total_net_profit = (df_val['market_cap'] / df_val['pe_ratio']).sum()
                            index_pe = total_market_cap / total_net_profit
                            indicators[f'{name}_PE'] = round(index_pe, 2)
                            
                            # 计算指数PB = 总市值 / 总净资产
                            # 其中：总净资产 = Σ(个股市值 / 个股PB)
                            total_net_assets = (df_val['market_cap'] / df_val['pb_ratio']).sum()
                            index_pb = total_market_cap / total_net_assets
                            indicators[f'{name}_PB'] = round(index_pb, 2)
                            
                            # 获取历史PE/PB用于计算分位数（简化处理：使用指数价格分位数作为代理）
                            df_index = self.fetcher.get_index_daily(code, period=252)
                            if df_index is not None and not df_index.empty:
                                current_price = df_index['收盘'].iloc[-1]
                                price_percentile = (df_index['收盘'] < current_price).mean() * 100
                                # 用价格分位数作为估值分位数的代理（价格越低估值越低）
                                indicators[f'{name}_PE_percentile'] = round(100 - price_percentile, 2)
                                indicators[f'{name}_PE_percentile_days'] = len(df_index)
                                indicators[f'{name}_PB_percentile'] = round(100 - price_percentile, 2)
                                indicators[f'{name}_PB_percentile_days'] = len(df_index)
                        
            except Exception as e:
                print(f"获取{name}估值数据失败: {e}")
        
        # 成长/价值比价
        try:
            df_cy = self.fetcher.get_index_daily("399006.XSHE", period=20)
            df_hs300 = self.fetcher.get_index_daily("000300.XSHG", period=20)
            
            if df_cy is not None and df_hs300 is not None:
                cy_return = (df_cy['收盘'].iloc[-1] / df_cy['收盘'].iloc[0] - 1) * 100
                hs300_return = (df_hs300['收盘'].iloc[-1] / df_hs300['收盘'].iloc[0] - 1) * 100
                indicators['growth_value_ratio'] = round(cy_return - hs300_return, 2)
        except Exception as e:
            print(f"计算成长价值比价失败: {e}")
        
        return indicators
    
    # ==================== 量化多头指标 ====================
    
    def get_market_divergence_indicators(self):
        """计算一九行情指标 —— 市值加权收益 vs 等权收益差值
        
        通过比较全A股票的市值加权日收益率与等权日收益率的差值，
        衡量市场收益是否集中在少数大市值股票上。
        差值为正且偏大时，说明大盘股涨、多数个股不涨，即一九行情。
        """
        indicators = {}
        
        try:
            end_date = dt.datetime.now().strftime('%Y-%m-%d')
            
            # 获取中证全指成分股作为全A代表
            stocks = get_index_stocks('000985.XSHG')
            if not stocks or len(stocks) == 0:
                print("获取中证全指成分股失败")
                return indicators
            
            # 获取所有成分股当日市值
            q = query(
                valuation.code,
                valuation.market_cap
            ).filter(
                valuation.market_cap != None,
                valuation.code.in_(stocks)
            )
            df_cap = get_fundamentals(q, end_date)
            
            if df_cap is None or df_cap.empty:
                print("获取市值数据失败")
                return indicators
            
            df_cap = df_cap.dropna(subset=['market_cap'])
            valid_stocks = df_cap['code'].tolist()
            
            if len(valid_stocks) == 0:
                return indicators
            
            # 获取近252+1个交易日的收盘价用于计算历史分位数
            # +1是因为计算日收益率需要前一天的数据
            history_period = 253
            df_prices = get_price(valid_stocks, count=history_period, end_date=end_date, frequency='daily', fields=['close'], panel=False)
            
            if df_prices is None or df_prices.empty:
                print("获取个股价格数据失败")
                return indicators
            
            # 透视为 日期×股票 的收盘价矩阵
            price_matrix = df_prices.pivot(index='time', columns='code', values='close')
            price_matrix = price_matrix.sort_index()
            
            # 计算日收益率
            returns_matrix = price_matrix.pct_change().dropna(how='all')
            
            if returns_matrix.empty:
                return indicators
            
            # 构建市值权重（使用最新市值作为权重，对历史序列保持一致）
            cap_dict = df_cap.set_index('code')['market_cap']
            common_stocks = returns_matrix.columns.intersection(cap_dict.index)
            returns_matrix = returns_matrix[common_stocks]
            weights = cap_dict[common_stocks]
            weights = weights / weights.sum()
            
            # 逐日计算市值加权收益和等权收益
            cap_weighted_returns = returns_matrix.multiply(weights, axis=1).sum(axis=1)
            # 等权收益：每日所有有效股票的平均收益
            equal_weighted_returns = returns_matrix.mean(axis=1)
            
            # 分化度 = 市值加权 - 等权（正值表示大盘股跑赢多数个股）
            divergence_series = (cap_weighted_returns - equal_weighted_returns) * 100
            
            # --- 当日分化度 ---
            latest_divergence = divergence_series.iloc[-1]
            indicators['market_divergence_daily'] = round(latest_divergence, 4)
            
            # --- 近20日累计分化度（滚动求和） ---
            if len(divergence_series) >= 20:
                divergence_20d = divergence_series.tail(20).sum()
                indicators['market_divergence_20d'] = round(divergence_20d, 4)
            
            # --- 分化度的历史分位数（用20日累计分化度的滚动窗口） ---
            if len(divergence_series) >= 40:
                rolling_20d = divergence_series.rolling(window=20).sum().dropna()
                percentile, days = self.calculate_percentile_with_days(rolling_20d, window=252, min_days=20)
                if percentile is not None:
                    indicators['market_divergence_percentile'] = percentile
                    indicators['market_divergence_percentile_days'] = days
            
            # 不再内置「一九指数」主观阈值打分；仅输出分化度序列，供自行配置 SCORE_WEIGHTS 后评分
            
        except Exception as e:
            print(f"计算一九行情指标失败: {e}")
        
        return indicators
    
    def get_quant_indicators(self):
        """获取量化多头指标"""
        indicators = {}
        
        # 1. 交投活跃度
        try:
            df_index = self.fetcher.get_index_daily("000001.XSHG", period=40)
            if df_index is not None and not df_index.empty:
                avg_amount_20 = df_index['成交额'].tail(20).mean() / 1e8
                indicators['market_amount_20d_avg'] = round(avg_amount_20, 2)
        except Exception as e:
            print(f"计算市场成交额失败: {e}")
        
        # 2. 指数拥挤度 - 使用中证全指成交额作为全市场A股总成交额
        index_symbols = {
            '沪深300': '000300.XSHG',
            '中证500': '000905.XSHG',
            '中证1000': '000852.XSHG',
        }
        
        try:
            # 获取中证全指成交额（包含沪深两市全部A股）
            df_all = self.fetcher.get_index_daily("000985.XSHG", period=20)  # 中证全指
            
            if df_all is not None and not df_all.empty:
                total_amount = df_all['成交额'].sum()
                
                for name, code in index_symbols.items():
                    df_idx = self.fetcher.get_index_daily(code, period=20)
                    if df_idx is not None and not df_idx.empty:
                        idx_amount = df_idx['成交额'].sum()
                        indicators[f'{name}_crowding'] = round(idx_amount / total_amount * 100, 2)
        except Exception as e:
            print(f"计算拥挤度失败: {e}")
        
        # 3. 一九行情分化指标
        try:
            divergence = self.get_market_divergence_indicators()
            indicators.update(divergence)
        except Exception as e:
            print(f"获取一九行情指标失败: {e}")
        
        return indicators
    
    # ==================== CTA策略指标（仅商品期货） ====================
    
    def get_cta_indicators(self):
        """获取CTA策略指标 - 仅商品期货"""
        indicators = {}
        
        # 商品期货品种列表（含黄金）
        commodity_symbols = ['RB', 'CU', 'SC', 'M', 'I', 'AU']  # AU=黄金
        
        # 1. 商品期货波动率及分位数
        try:
            volatilities = []
            for symbol in commodity_symbols:
                try:
                    df = self.fetcher.get_futures_main_contract(symbol, period=252)
                    if df is not None and not df.empty and len(df) >= 20:
                        df = df.sort_values('日期')
                        df['daily_return'] = df['收盘价'].pct_change()
                        
                        # 计算20日滚动波动率序列
                        df['volatility_20d'] = df['daily_return'].rolling(20).std() * np.sqrt(252) * 100
                        
                        # 当前波动率
                        current_vol = df['volatility_20d'].iloc[-1]
                        if not np.isnan(current_vol):
                            indicators[f'{symbol}_volatility_20d'] = round(current_vol, 2)
                            volatilities.append(current_vol)
                        
                        # 波动率近一年分位数（记录实际天数）
                        vol_percentile, vol_days = self.calculate_percentile_with_days(df['volatility_20d'], window=252, min_days=20)
                        if vol_percentile is not None:
                            indicators[f'{symbol}_volatility_percentile'] = vol_percentile
                            indicators[f'{symbol}_volatility_percentile_days'] = vol_days
                except:
                    continue
            
            if volatilities:
                indicators['commodity_avg_volatility'] = round(np.mean(volatilities), 2)
        except Exception as e:
            print(f"计算商品期货波动率失败: {e}")
        
        # 2. 商品期货成交量及分位数
        try:
            total_volumes = []
            for symbol in commodity_symbols:
                try:
                    df = self.fetcher.get_futures_main_contract(symbol, period=252)
                    if df is not None and not df.empty:
                        # 近20日平均成交量
                        avg_vol = df['成交量'].tail(20).mean()
                        if not np.isnan(avg_vol):
                            indicators[f'{symbol}_avg_volume'] = round(avg_vol, 2)
                            total_volumes.append(avg_vol)
                        
                        # 成交量近一年分位数（记录实际天数）
                        vol_percentile, vol_days = self.calculate_percentile_with_days(df['成交量'], window=252, min_days=20)
                        if vol_percentile is not None:
                            indicators[f'{symbol}_volume_percentile'] = vol_percentile
                            indicators[f'{symbol}_volume_percentile_days'] = vol_days
                except:
                    continue
            
            if total_volumes:
                indicators['commodity_total_volume'] = round(np.sum(total_volumes), 2)
        except Exception as e:
            print(f"计算商品期货成交量失败: {e}")
        
        # 3. 商品期货持仓量及分位数
        try:
            total_positions = 0
            for symbol in commodity_symbols:
                try:
                    df_pos = self.fetcher.get_futures_info_data(symbol)
                    if df_pos is not None and not df_pos.empty:
                        # 最新持仓量
                        latest_pos = df_pos.iloc[-1].values[0] if len(df_pos.columns) > 0 else 0
                        indicators[f'{symbol}_positions'] = round(latest_pos, 2)
                        total_positions += latest_pos
                        
                        # 持仓量近一年分位数（记录实际天数）
                        pos_series = df_pos.iloc[:, 0] if len(df_pos.columns) > 0 else pd.Series()
                        pos_percentile, pos_days = self.calculate_percentile_with_days(pos_series, window=252, min_days=20)
                        if pos_percentile is not None:
                            indicators[f'{symbol}_positions_percentile'] = pos_percentile
                            indicators[f'{symbol}_positions_percentile_days'] = pos_days
                except:
                    continue
            
            indicators['commodity_total_positions'] = round(total_positions, 2)
        except Exception as e:
            print(f"计算商品期货持仓量失败: {e}")
        
        return indicators
    
    # ==================== 套利策略指标 ====================
    
    def get_arbitrage_indicators(self):
        """获取套利策略指标 - ETF套利/股指高频套利/期权套利/基差"""
        indicators = {}
        
        # ===== 1. ETF套利指标（主流指数ETF） =====
        # 主流指数ETF列表
        etf_list = {
            '510300.XSHG': '300ETF',    # 沪深300ETF
            '510500.XSHG': '500ETF',    # 中证500ETF
            '512100.XSHG': '1000ETF',   # 中证1000ETF
            '510050.XSHG': '50ETF',     # 上证50ETF
            '159915.XSHE': '创业板ETF', # 创业板ETF
            '588000.XSHG': '科创50ETF', # 科创50ETF
        }
        
        # 对应的指数代码
        index_map = {
            '510300.XSHG': '000300.XSHG',
            '510500.XSHG': '000905.XSHG',
            '512100.XSHG': '000852.XSHG',
            '510050.XSHG': '000016.XSHG',
            '159915.XSHE': '399006.XSHE',
            '588000.XSHG': '000688.XSHG',
        }
        
        for etf_code, etf_name in etf_list.items():
            try:
                # 获取ETF数据（近一年）
                df_etf = self.fetcher.get_etf_daily(etf_code, period=252)
                df_nav = self.fetcher.get_etf_nav(etf_code, period=252)
                
                if df_etf is not None and not df_etf.empty:
                    # 1.1 ETF成交额及分位数（记录实际天数）
                    avg_amount_20d = df_etf['成交额'].tail(20).mean() / 1e8
                    indicators[f'{etf_name}_avg_amount_20d'] = round(avg_amount_20d, 2)
                    
                    amount_percentile, amount_days = self.calculate_percentile_with_days(df_etf['成交额'], window=252, min_days=20)
                    if amount_percentile is not None:
                        indicators[f'{etf_name}_amount_percentile'] = amount_percentile
                        indicators[f'{etf_name}_amount_percentile_days'] = amount_days
                    
                    # 1.2 ETF波动率及分位数（记录实际天数）
                    df_etf = df_etf.sort_values('日期')
                    df_etf['daily_return'] = df_etf['收盘'].pct_change()
                    df_etf['volatility_20d'] = df_etf['daily_return'].rolling(20).std() * np.sqrt(252) * 100
                    
                    current_vol = df_etf['volatility_20d'].iloc[-1]
                    if not np.isnan(current_vol):
                        indicators[f'{etf_name}_volatility_20d'] = round(current_vol, 2)
                    
                    vol_percentile, vol_days = self.calculate_percentile_with_days(df_etf['volatility_20d'], window=252, min_days=20)
                    if vol_percentile is not None:
                        indicators[f'{etf_name}_volatility_percentile'] = vol_percentile
                        indicators[f'{etf_name}_volatility_percentile_days'] = vol_days
                    
                    # 1.3 ETF折溢价率及分位数 - 使用单位净值(NAV)计算
                    if df_nav is not None and not df_nav.empty:
                        # 合并ETF价格和净值数据
                        df_premium = pd.merge(df_etf[['日期', '收盘']], df_nav[['日期', '单位净值']], 
                                             on='日期', how='inner')
                        
                        if len(df_premium) > 0:
                            # 计算折溢价率 = (收盘价 / 单位净值 - 1) * 100%
                            df_premium['premium'] = (df_premium['收盘'] / df_premium['单位净值'] - 1) * 100
                            
                            # 当前折溢价率
                            current_premium = df_premium['premium'].iloc[-1]
                            indicators[f'{etf_name}_premium'] = round(current_premium, 4)
                            
                            # 折溢价率近一年分位数（记录实际天数）
                            premium_percentile, premium_days = self.calculate_percentile_with_days(df_premium['premium'], window=252, min_days=20)
                            if premium_percentile is not None:
                                indicators[f'{etf_name}_premium_percentile'] = premium_percentile
                                indicators[f'{etf_name}_premium_percentile_days'] = premium_days
            except Exception as e:
                print(f"计算{etf_name}指标失败: {e}")
        
        # ===== 2. 股指高频套利指标 =====
        # 股指期货近20日波动率、波动率分位数、成交额分位数
        futures_map = {
            'IF': '沪深300',
            'IC': '中证500',
            'IM': '中证1000'
        }
        
        for future_code, index_name in futures_map.items():
            try:
                df_future = self.fetcher.get_futures_main_contract(future_code, period=252)
                if df_future is not None and not df_future.empty and len(df_future) >= 20:
                    df_future = df_future.sort_values('日期')
                    df_future['daily_return'] = df_future['收盘价'].pct_change()
                    
                    # 计算20日滚动波动率序列
                    df_future['volatility_20d'] = df_future['daily_return'].rolling(20).std() * np.sqrt(252) * 100
                    
                    # 当前波动率
                    current_vol = df_future['volatility_20d'].iloc[-1]
                    if not np.isnan(current_vol):
                        indicators[f'{future_code}_volatility_20d'] = round(current_vol, 2)
                    
                    # 波动率近一年分位数（记录实际天数）
                    vol_percentile, vol_days = self.calculate_percentile_with_days(df_future['volatility_20d'], window=252, min_days=20)
                    if vol_percentile is not None:
                        indicators[f'{future_code}_volatility_percentile'] = vol_percentile
                        indicators[f'{future_code}_volatility_percentile_days'] = vol_days
                    
                    # 成交额近一年分位数（记录实际天数）
                    amount_percentile, amount_days = self.calculate_percentile_with_days(df_future['成交额'], window=252, min_days=20)
                    if amount_percentile is not None:
                        indicators[f'{future_code}_amount_percentile'] = amount_percentile
                        indicators[f'{future_code}_amount_percentile_days'] = amount_days
            except:
                continue
        
        # ===== 3. 期权套利指标 =====
        # 动态获取期权合约的隐含波动率数据
        
        def get_option_iv(underlying_code, option_type='ETF', name=''):
            """获取期权的隐含波动率 - 扩大扫描范围"""
            try:
                end_date = dt.datetime.now().strftime('%Y-%m-%d')
                
                # 使用更大的扫描范围
                if option_type == 'ETF':
                    # ETF期权代码范围：10000000-10010000 (沪市)
                    option_list = []
                    # 尝试多个代码段
                    code_ranges = [
                        range(10002760, 10002900),
                        range(10004000, 10004500),
                        range(10005000, 10005500),
                    ]
                    
                    for code_range in code_ranges:
                        for i in code_range:
                            try:
                                code = f'{i}.XSHG'
                                df = get_price(code, count=3, end_date=end_date, frequency='daily', fields=['close', 'implied_volatility'])
                                if df is not None and not df.empty:
                                    last_iv = df['implied_volatility'].iloc[-1]
                                    if pd.notna(last_iv) and last_iv > 0:
                                        option_list.append((code, last_iv))
                            except:
                                continue
                        if len(option_list) >= 3:  # 找到足够合约就停止
                            break
                    
                    if option_list:
                        # 取IV中位数的合约（避免极端值）
                        option_list.sort(key=lambda x: x[1])
                        best_code = option_list[len(option_list)//2][0]
                        
                        # 获取该合约的历史IV
                        df_option = get_price(best_code, count=252, end_date=end_date, frequency='daily', fields=['close', 'implied_volatility'])
                        if df_option is not None and not df_option.empty:
                            df_option = df_option.reset_index()
                            df_option.rename(columns={'index': '日期', 'implied_volatility': 'IV'}, inplace=True)
                            df_option['日期'] = pd.to_datetime(df_option['日期'])
                            df_option = df_option.sort_values('日期')
                            df_option = df_option.dropna(subset=['IV'])
                            
                            if len(df_option) >= 20:
                                current_iv = df_option['IV'].iloc[-1]
                                indicators[f'{name}_IV'] = round(current_iv * 100, 2)
                                
                                iv_percentile, iv_days = self.calculate_percentile_with_days(
                                    df_option['IV'], window=252, min_days=20)
                                if iv_percentile is not None:
                                    indicators[f'{name}_IV_percentile'] = iv_percentile
                                    indicators[f'{name}_IV_percentile_days'] = iv_days
                                    return True
                
                elif option_type == 'INDEX':
                    # 股指期权代码范围
                    option_list = []
                    code_ranges = [
                        range(10002764, 10002900),
                        range(10006000, 10006500),
                    ]
                    
                    for code_range in code_ranges:
                        for i in code_range:
                            try:
                                code = f'{i}.XSHG'
                                df = get_price(code, count=3, end_date=end_date, frequency='daily', fields=['close', 'implied_volatility'])
                                if df is not None and not df.empty:
                                    last_iv = df['implied_volatility'].iloc[-1]
                                    if pd.notna(last_iv) and last_iv > 0:
                                        option_list.append((code, last_iv))
                            except:
                                continue
                        if len(option_list) >= 2:
                            break
                    
                    if option_list:
                        option_list.sort(key=lambda x: x[1])
                        best_code = option_list[len(option_list)//2][0]
                        
                        df_option = get_price(best_code, count=252, end_date=end_date, frequency='daily', fields=['close', 'implied_volatility'])
                        if df_option is not None and not df_option.empty:
                            df_option = df_option.reset_index()
                            df_option.rename(columns={'index': '日期', 'implied_volatility': 'IV'}, inplace=True)
                            df_option['日期'] = pd.to_datetime(df_option['日期'])
                            df_option = df_option.sort_values('日期')
                            df_option = df_option.dropna(subset=['IV'])
                            
                            if len(df_option) > 0:
                                current_iv = df_option['IV'].iloc[-1]
                                indicators[f'{name}_IV'] = round(current_iv * 100, 2)
                                
                                iv_percentile, iv_days = self.calculate_percentile_with_days(
                                    df_option['IV'], window=252, min_days=20)
                                if iv_percentile is not None:
                                    indicators[f'{name}_IV_percentile'] = iv_percentile
                                    indicators[f'{name}_IV_percentile_days'] = iv_days
                                    return True
                
                return False
            except Exception as e:
                print(f"获取{name}期权IV失败: {e}")
                return False
        
        # 尝试获取各类期权IV
        get_option_iv('510050.XSHG', 'ETF', '50ETF')
        get_option_iv('510300.XSHG', 'ETF', '300ETF')
        get_option_iv('510500.XSHG', 'ETF', '500ETF')
        get_option_iv('588000.XSHG', 'ETF', '科创50ETF')
        get_option_iv('159915.XSHE', 'ETF', '创业板ETF')
        
        # 股指期权
        get_option_iv('000300.XSHG', 'INDEX', '300股指')
        get_option_iv('000852.XSHG', 'INDEX', '1000股指')
        get_option_iv('000016.XSHG', 'INDEX', '50股指')
        
        # ===== 4. 股指期货基差 =====
        futures_basis_map = {
            'IF': '000300.XSHG',
            'IC': '000905.XSHG',
            'IM': '000852.XSHG'
        }
        
        for future_code, index_code in futures_basis_map.items():
            try:
                df_future = self.fetcher.get_futures_main_contract(future_code, period=5)
                df_index = self.fetcher.get_index_daily(index_code, period=5)
                
                if df_future is not None and df_index is not None:
                    future_price = df_future['收盘价'].iloc[-1]
                    spot_price = df_index['收盘'].iloc[-1]
                    
                    if future_price and spot_price:
                        basis = (future_price - spot_price) / spot_price * 100
                        indicators[f'{future_code}_basis'] = round(basis, 4)
                        basis_annual = basis * 4
                        indicators[f'{future_code}_basis_annual'] = round(basis_annual, 2)
            except Exception as e:
                print(f"计算{future_code}基差失败: {e}")
        
        return indicators
    
    # ==================== 市场中性策略指标 ====================
    
    def get_market_neutral_indicators(self):
        """获取市场中性策略指标"""
        indicators = {}
        
        # 1. 交投活跃度
        try:
            df_index = self.fetcher.get_index_daily("000001.XSHG", period=40)
            if df_index is not None and not df_index.empty:
                avg_amount_20 = df_index['成交额'].tail(20).mean() / 1e8
                indicators['market_amount_20d_avg'] = round(avg_amount_20, 2)
        except Exception as e:
            print(f"计算市场成交额失败: {e}")
        
        # 2. 大中小盘拥挤度 - 使用中证全指成交额作为全市场A股总成交额
        index_symbols = {
            '沪深300': '000300.XSHG',
            '中证500': '000905.XSHG',
            '中证1000': '000852.XSHG',
        }
        
        try:
            # 获取中证全指成交额（包含沪深两市全部A股）
            df_all = self.fetcher.get_index_daily("000985.XSHG", period=20)  # 中证全指
            
            if df_all is not None and not df_all.empty:
                total_amount = df_all['成交额'].sum()
                
                for name, code in index_symbols.items():
                    df_idx = self.fetcher.get_index_daily(code, period=20)
                    if df_idx is not None and not df_idx.empty:
                        idx_amount = df_idx['成交额'].sum()
                        indicators[f'{name}_crowding'] = round(idx_amount / total_amount * 100, 2)
        except Exception as e:
            print(f"计算拥挤度失败: {e}")
        
        # 3. 股指期货基差
        futures_map = {
            'IF': '000300.XSHG',
            'IC': '000905.XSHG',
            'IM': '000852.XSHG'
        }
        
        for future_code, index_code in futures_map.items():
            try:
                df_future = self.fetcher.get_futures_main_contract(future_code, period=5)
                df_index = self.fetcher.get_index_daily(index_code, period=5)
                
                if df_future is not None and df_index is not None:
                    future_price = df_future['收盘价'].iloc[-1]
                    spot_price = df_index['收盘'].iloc[-1]
                    
                    if future_price and spot_price:
                        basis = (future_price - spot_price) / spot_price * 100
                        indicators[f'{future_code}_basis'] = round(basis, 4)
                        indicators[f'{future_code}_basis_annual'] = round(basis * 4, 2)
            except Exception as e:
                print(f"计算{future_code}基差失败: {e}")
        
        # 4. 市场形态指标（哑铃型/纺锤型）
        try:
            # 获取三个指数的日线数据（需要至少6天数据计算5日均线）
            df_2000 = self.fetcher.get_index_daily("399303.XSHE", period=10)  # 国证2000
            df_300 = self.fetcher.get_index_daily("000300.XSHG", period=10)   # 沪深300
            df_500 = self.fetcher.get_index_daily("000905.XSHG", period=10)   # 中证500
            
            # 检查数据是否成功获取
            if df_2000 is None:
                print("获取国证2000数据失败，无法计算市场形态指标")
            else:
                print(f"国证2000数据量：{len(df_2000)}天")
            if df_300 is None:
                print("获取沪深300数据失败，无法计算市场形态指标")
            else:
                print(f"沪深300数据量：{len(df_300)}天")
            if df_500 is None:
                print("获取中证500数据失败，无法计算市场形态指标")
            else:
                print(f"中证500数据量：{len(df_500)}天")
            
            if df_2000 is not None and df_300 is not None and df_500 is not None:
                # 设置日期为索引
                df_2000 = df_2000.set_index('日期')
                df_300 = df_300.set_index('日期')
                df_500 = df_500.set_index('日期')
                
                # 计算日收益率
                df_2000['return'] = df_2000['收盘'].pct_change() * 100
                df_300['return'] = df_300['收盘'].pct_change() * 100
                df_500['return'] = df_500['收盘'].pct_change() * 100
                
                # 合并数据（使用join确保日期对齐）
                df_merged = pd.DataFrame({
                    'r2000': df_2000['return'],
                    'r300': df_300['return'],
                    'r500': df_500['return']
                }).dropna()
                
                print(f"市场形态指标：合并后数据量={len(df_merged)}天")
                
                if len(df_merged) >= 6:
                    # 计算哑铃指数 = 国证2000 + 沪深300 - 中证500
                    df_merged['dumbbell'] = df_merged['r2000'] + df_merged['r300'] - df_merged['r500']
                    
                    # 计算5日均线
                    df_merged['ma5'] = df_merged['dumbbell'].rolling(5).mean()
                    
                    # 获取当前值和5日均线值
                    current_dumbbell = df_merged['dumbbell'].iloc[-1]
                    current_ma5 = df_merged['ma5'].iloc[-1]
                    prev_dumbbell = df_merged['dumbbell'].iloc[-2]
                    prev_ma5 = df_merged['ma5'].iloc[-2]
                    
                    # 存储指标值
                    indicators['dumbbell_index'] = round(current_dumbbell, 4)
                    indicators['dumbbell_ma5'] = round(current_ma5, 4)
                    
                    # 判断形态变化
                    # 向上突破5日均线：当前值>MA5 且 前值<=前MA5
                    if current_dumbbell > current_ma5 and prev_dumbbell <= prev_ma5:
                        indicators['market_pattern'] = '哑铃型'
                        indicators['pattern_signal'] = 1  # 向上突破
                        print(f"市场形态：哑铃型（哑铃指数={current_dumbbell:.4f}%，MA5={current_ma5:.4f}%）")
                    # 向下突破5日均线：当前值<MA5 且 前值>=前MA5
                    elif current_dumbbell < current_ma5 and prev_dumbbell >= prev_ma5:
                        indicators['market_pattern'] = '纺锤型'
                        indicators['pattern_signal'] = -1  # 向下突破
                        print(f"市场形态：纺锤型（哑铃指数={current_dumbbell:.4f}%，MA5={current_ma5:.4f}%）")
                    else:
                        indicators['market_pattern'] = '无形态变化'
                        indicators['pattern_signal'] = 0  # 无变化
                        print(f"市场形态：无形态变化（哑铃指数={current_dumbbell:.4f}%，MA5={current_ma5:.4f}%）")
                else:
                    print(f"市场形态指标：数据不足，合并后只有{len(df_merged)}天，需要至少6天")
        except Exception as e:
            print(f"计算市场形态指标失败: {e}")
        
        return indicators


# ============================================
# 策略评分模块
# ============================================

class StrategyScorer:
    """策略评分类：不在代码中写死阈值与加减分；请配置全局 SCORE_WEIGHTS 后扩展 _score_from_weights。"""
    
    def __init__(self):
        self.calculator = IndicatorCalculator()

    def score_subjective_long(self):
        score, details = 50, []
        if SCORE_WEIGHTS.get('subjective'):
            details.append("已配置 subjective 权重但未实现自动解析，请自行实现或保持中性 50 分")
        else:
            details.append("未配置 SCORE_WEIGHTS['subjective']：不代入默认成交额/PE 等占位，评分保持 50")
        return {'score': score, 'details': details, 'level': self._get_level(score), 'trend': '→ 持平'}

    def score_quant_long(self):
        score, details = 50, []
        if SCORE_WEIGHTS.get('quant'):
            details.append("已配置 quant 权重但未实现自动解析，请自行实现或保持中性 50 分")
        else:
            details.append("未配置 SCORE_WEIGHTS['quant']：不使用拥挤度/一九指数等硬编码阈值，评分保持 50")
        return {'score': score, 'details': details, 'level': self._get_level(score), 'trend': '→ 持平'}

    def score_cta(self):
        score, details = 50, []
        if SCORE_WEIGHTS.get('cta'):
            details.append("已配置 cta 权重但未实现自动解析，请自行实现或保持中性 50 分")
        else:
            details.append("未配置 SCORE_WEIGHTS['cta']：不使用分位数硬编码阈值，评分保持 50")
        return {'score': score, 'details': details, 'level': self._get_level(score), 'trend': '→ 持平'}

    def score_etf_arbitrage(self):
        score, details = 50, []
        if SCORE_WEIGHTS.get('etf'):
            details.append("已配置 etf 权重但未实现自动解析，请自行实现或保持中性 50 分")
        else:
            details.append("未配置 SCORE_WEIGHTS['etf']：不使用波动率/成交额分位等硬编码，评分保持 50")
        return {'score': score, 'details': details, 'level': self._get_level(score), 'trend': '→ 持平'}

    def score_index_arbitrage(self):
        score, details = 50, []
        if SCORE_WEIGHTS.get('index_arb'):
            details.append("已配置 index_arb 权重但未实现自动解析，请自行实现或保持中性 50 分")
        else:
            details.append("未配置 SCORE_WEIGHTS['index_arb']：评分保持 50")
        return {'score': score, 'details': details, 'level': self._get_level(score), 'trend': '→ 持平'}

    def score_option_arbitrage(self):
        score, details = 50, []
        if SCORE_WEIGHTS.get('option'):
            details.append("已配置 option 权重但未实现自动解析，请自行实现或保持中性 50 分")
        else:
            details.append("未配置 SCORE_WEIGHTS['option']：评分保持 50")
        return {'score': score, 'details': details, 'level': self._get_level(score), 'trend': '→ 持平'}

    def score_market_neutral(self):
        score, details = 50, []
        if SCORE_WEIGHTS.get('neutral'):
            details.append("已配置 neutral 权重但未实现自动解析，请自行实现或保持中性 50 分")
        else:
            details.append("未配置 SCORE_WEIGHTS['neutral']：不使用基差/拥挤度硬编码档，评分保持 50")
        return {'score': score, 'details': details, 'level': self._get_level(score), 'trend': '→ 持平'}
    
    def _get_level(self, score):
        """根据分数获取适配环境等级"""
        if score >= 80:
            return "非常适宜"
        elif score >= 65:
            return "偏多"
        elif score >= 45:
            return "中性"
        elif score >= 30:
            return "中性偏空"
        else:
            return "偏空"
    
    def get_all_scores(self):
        """获取所有策略评分"""
        return {
            '主观多头': self.score_subjective_long(),
            '量化多头': self.score_quant_long(),
            'CTA策略': self.score_cta(),
            'ETF套利': self.score_etf_arbitrage(),
            '股指套利': self.score_index_arbitrage(),
            '期权套利': self.score_option_arbitrage(),
            '市场中性': self.score_market_neutral()
        }


# ============================================
# Excel报告生成模块
# ============================================

class ExcelReportGenerator:
    """Excel报告生成类"""
    
    def __init__(self, output_path=None):
        if output_path is None:
            timestamp = dt.datetime.now().strftime('%Y%m%d_%H%M%S')
            # 聚宽环境中使用当前目录
            self.output_path = f'strategy_monitor_report_{timestamp}.xlsx'
        else:
            self.output_path = output_path
    
    def generate_excel_report(self, all_scores, all_indicators):
        """生成Excel报告"""
        wb = Workbook()
        
        if 'Sheet' in wb.sheetnames:
            wb.remove(wb['Sheet'])
        
        self._create_overview_sheet(wb, all_scores)
        self._create_subjective_sheet(wb, all_indicators.get('subjective', {}))
        self._create_quant_sheet(wb, all_indicators.get('quant', {}))
        self._create_cta_sheet(wb, all_indicators.get('cta', {}))
        self._create_arbitrage_sheet(wb, all_indicators.get('arbitrage', {}))
        self._create_neutral_sheet(wb, all_indicators.get('neutral', {}))
        self._create_score_details_sheet(wb, all_scores)
        
        wb.save(self.output_path)
        print(f"Excel报告已保存至: {self.output_path}")
        return self.output_path
    
    def _create_overview_sheet(self, wb, all_scores):
        """创建总览sheet"""
        ws = wb.create_sheet("策略评分总览", 0)
        
        ws['A1'] = '策略环境监测报告'
        ws['A1'].font = Font(bold=True, size=16, color='1F4E78')
        ws.merge_cells('A1:E1')
        ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
        
        ws['A2'] = f'生成时间: {dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")}'
        ws['A2'].font = Font(size=10, color='666666')
        ws.merge_cells('A2:E2')
        
        headers = ['策略类型', '评分', '适配环境', '趋势', '评分说明']
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=4, column=col, value=header)
            cell.font = Font(bold=True, color='FFFFFF')
            cell.fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = Border(
                left=Side(style='thin'), right=Side(style='thin'),
                top=Side(style='thin'), bottom=Side(style='thin')
            )
        
        row = 5
        for strategy_name, result in all_scores.items():
            ws.cell(row=row, column=1, value=strategy_name)
            ws.cell(row=row, column=2, value=result['score'])
            ws.cell(row=row, column=3, value=result['level'])
            ws.cell(row=row, column=4, value=result['trend'])
            
            details = result.get('details', [])
            detail_text = details[0] if details else ''
            ws.cell(row=row, column=5, value=detail_text)
            
            score = result['score']
            if score >= 80:
                fill_color = 'C6EFCE'
            elif score >= 65:
                fill_color = 'E2EFDA'
            elif score >= 45:
                fill_color = 'FFEB9C'
            elif score >= 30:
                fill_color = 'FFC7CE'
            else:
                fill_color = 'FF6B6B'
            
            for col in range(1, 6):
                cell = ws.cell(row=row, column=col)
                cell.fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type='solid')
                cell.border = Border(
                    left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin')
                )
                cell.alignment = Alignment(horizontal='center' if col <= 4 else 'left', vertical='center')
            
            row += 1
        
        ws.column_dimensions['A'].width = 15
        ws.column_dimensions['B'].width = 10
        ws.column_dimensions['C'].width = 12
        ws.column_dimensions['D'].width = 10
        ws.column_dimensions['E'].width = 40
    
    def _create_subjective_sheet(self, wb, indicators):
        """创建主观多头指标sheet"""
        ws = wb.create_sheet("主观多头指标")
        
        ws['A1'] = '主观多头策略 - 指标详情'
        ws['A1'].font = Font(bold=True, size=14, color='1F4E78')
        ws.merge_cells('A1:D1')
        ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
        
        headers = ['指标名称', '当前值', '单位', '指标说明']
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=3, column=col, value=header)
            cell.font = Font(bold=True, color='FFFFFF')
            cell.fill = PatternFill(start_color='70AD47', end_color='70AD47', fill_type='solid')
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = Border(
                left=Side(style='thin'), right=Side(style='thin'),
                top=Side(style='thin'), bottom=Side(style='thin')
            )
        
        row = 4
        
        # ===== 第一部分：主要指数当日涨跌幅 =====
        ws.cell(row=row, column=1, value='【主要指数当日涨跌幅】')
        ws.cell(row=row, column=1).font = Font(bold=True, size=11, color='1F4E78')
        ws.merge_cells(f'A{row}:D{row}')
        row += 1
        
        index_list = ['沪深300', '中证A500', '中证500', '中证1000', '国证2000']
        for idx_name in index_list:
            ret = indicators.get(f'{idx_name}_daily_return')
            if ret is not None:
                ws.cell(row=row, column=1, value=f'{idx_name}当日涨跌幅')
                ws.cell(row=row, column=2, value=f"{ret}%")
                ws.cell(row=row, column=3, value='%')
                ws.cell(row=row, column=4, value=f'{idx_name}指数当日涨跌幅')
                for col in range(1, 5):
                    cell = ws.cell(row=row, column=col)
                    cell.border = Border(left=Side(style='thin'), right=Side(style='thin'),
                                        top=Side(style='thin'), bottom=Side(style='thin'))
                    cell.alignment = Alignment(horizontal='left' if col in [1, 4] else 'center', vertical='center')
                row += 1
        
        row += 1  # 空行
        
        # ===== 第二部分：市场活跃度 =====
        ws.cell(row=row, column=1, value='【市场活跃度】')
        ws.cell(row=row, column=1).font = Font(bold=True, size=11, color='1F4E78')
        ws.merge_cells(f'A{row}:D{row}')
        row += 1
        
        market_amount = indicators.get('market_amount_20d_avg')
        if market_amount is not None:
            ws.cell(row=row, column=1, value='近20日市场平均成交额')
            ws.cell(row=row, column=2, value=f"{market_amount:.2f}")
            ws.cell(row=row, column=3, value='亿元')
            ws.cell(row=row, column=4, value='反映市场整体活跃度')
            for col in range(1, 5):
                cell = ws.cell(row=row, column=col)
                cell.border = Border(left=Side(style='thin'), right=Side(style='thin'),
                                    top=Side(style='thin'), bottom=Side(style='thin'))
                cell.alignment = Alignment(horizontal='left' if col in [1, 4] else 'center', vertical='center')
            row += 1
        
        amount_change = indicators.get('market_amount_change_20d')
        if amount_change is not None:
            ws.cell(row=row, column=1, value='近20日成交额变化率')
            ws.cell(row=row, column=2, value=f"{amount_change}%")
            ws.cell(row=row, column=3, value='%')
            ws.cell(row=row, column=4, value='反映市场热度变化')
            for col in range(1, 5):
                cell = ws.cell(row=row, column=col)
                cell.border = Border(left=Side(style='thin'), right=Side(style='thin'),
                                    top=Side(style='thin'), bottom=Side(style='thin'))
                cell.alignment = Alignment(horizontal='left' if col in [1, 4] else 'center', vertical='center')
            row += 1
        
        row += 1  # 空行
        
        # ===== 第三部分：估值指标（PE/PB） =====
        ws.cell(row=row, column=1, value='【估值指标 - PE/PB分位数】')
        ws.cell(row=row, column=1).font = Font(bold=True, size=11, color='1F4E78')
        ws.merge_cells(f'A{row}:D{row}')
        row += 1
        
        index_valuation = [
            ('沪深300', '沪深300指数'),
            ('中证500', '中证500指数'),
            ('中证1000', '中证1000指数'),
            ('国证2000', '国证2000指数（小盘）')
        ]
        
        for idx_code, idx_name in index_valuation:
            pe = indicators.get(f'{idx_code}_PE')
            pe_pct = indicators.get(f'{idx_code}_PE_percentile')
            pe_days = indicators.get(f'{idx_code}_PE_percentile_days', 0)
            pb = indicators.get(f'{idx_code}_PB')
            pb_pct = indicators.get(f'{idx_code}_PB_percentile')
            pb_days = indicators.get(f'{idx_code}_PB_percentile_days', 0)
            
            if pe is not None:
                ws.cell(row=row, column=1, value=f'{idx_name} PE')
                ws.cell(row=row, column=2, value=f"{pe:.2f}倍")
                if pe_pct is not None:
                    pct_display = f"{pe_pct}% (基于{pe_days}天)" if pe_days > 0 else f"{pe_pct}%"
                    ws.cell(row=row, column=3, value=pct_display)
                else:
                    ws.cell(row=row, column=3, value='数据不足')
                ws.cell(row=row, column=4, value=f'{idx_name}市盈率及历史分位数')
                for col in range(1, 5):
                    cell = ws.cell(row=row, column=col)
                    cell.border = Border(left=Side(style='thin'), right=Side(style='thin'),
                                        top=Side(style='thin'), bottom=Side(style='thin'))
                    cell.alignment = Alignment(horizontal='left' if col in [1, 4] else 'center', vertical='center')
                row += 1
            
            if pb is not None:
                ws.cell(row=row, column=1, value=f'{idx_name} PB')
                ws.cell(row=row, column=2, value=f"{pb:.2f}倍")
                if pb_pct is not None:
                    pct_display = f"{pb_pct}% (基于{pb_days}天)" if pb_days > 0 else f"{pb_pct}%"
                    ws.cell(row=row, column=3, value=pct_display)
                else:
                    ws.cell(row=row, column=3, value='数据不足')
                ws.cell(row=row, column=4, value=f'{idx_name}市净率及历史分位数')
                for col in range(1, 5):
                    cell = ws.cell(row=row, column=col)
                    cell.border = Border(left=Side(style='thin'), right=Side(style='thin'),
                                        top=Side(style='thin'), bottom=Side(style='thin'))
                    cell.alignment = Alignment(horizontal='left' if col in [1, 4] else 'center', vertical='center')
                row += 1
        
        row += 1  # 空行
        
        # ===== 第四部分：风格指标 =====
        ws.cell(row=row, column=1, value='【风格指标】')
        ws.cell(row=row, column=1).font = Font(bold=True, size=11, color='1F4E78')
        ws.merge_cells(f'A{row}:D{row}')
        row += 1
        
        gv_ratio = indicators.get('growth_value_ratio')
        if gv_ratio is not None:
            ws.cell(row=row, column=1, value='成长/价值比价')
            ws.cell(row=row, column=2, value=f"{gv_ratio}%")
            ws.cell(row=row, column=3, value='%')
            ws.cell(row=row, column=4, value='创业板指与沪深300指数20日收益率差值，反映风格偏好')
            for col in range(1, 5):
                cell = ws.cell(row=row, column=col)
                cell.border = Border(left=Side(style='thin'), right=Side(style='thin'),
                                    top=Side(style='thin'), bottom=Side(style='thin'))
                cell.alignment = Alignment(horizontal='left' if col in [1, 4] else 'center', vertical='center')
            row += 1
        
        # 设置列宽
        ws.column_dimensions['A'].width = 24
        ws.column_dimensions['B'].width = 18
        ws.column_dimensions['C'].width = 22
        ws.column_dimensions['D'].width = 40
    
    def _create_quant_sheet(self, wb, indicators):
        """创建量化多头指标sheet"""
        ws = wb.create_sheet("量化多头指标")
        
        ws['A1'] = '量化多头策略 - 指标详情'
        ws['A1'].font = Font(bold=True, size=14, color='1F4E78')
        ws.merge_cells('A1:D1')
        ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
        
        headers = ['指标名称', '当前值', '单位', '指标说明']
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=3, column=col, value=header)
            cell.font = Font(bold=True, color='FFFFFF')
            cell.fill = PatternFill(start_color='5B9BD5', end_color='5B9BD5', fill_type='solid')
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = Border(
                left=Side(style='thin'), right=Side(style='thin'),
                top=Side(style='thin'), bottom=Side(style='thin')
            )
        
        indicator_descriptions = {
            'market_amount_20d_avg': '近20日市场平均成交额，反映市场活跃度',
            '沪深300_crowding': '沪深300指数成交额占全市场比例（近20日），反映大盘拥挤度',
            '中证500_crowding': '中证500指数成交额占全市场比例（近20日），反映中盘拥挤度',
            '中证1000_crowding': '中证1000指数成交额占全市场比例（近20日），反映小盘拥挤度',
            'market_divergence_daily': '当日市值加权收益与等权收益的差值，正值表示大盘股跑赢多数个股',
            'market_divergence_20d': '近20日分化度累计值，持续正值表示一九行情持续',
            'market_divergence_percentile': '20日累计分化度在近一年中的历史分位数',
            'market_divergence_percentile_days': '计算分化度分位数所用的历史天数',
        }
        
        indicator_units = {
            'market_amount_20d_avg': '亿元',
            '沪深300_crowding': '%',
            '中证500_crowding': '%',
            '中证1000_crowding': '%',
            'market_divergence_daily': '%',
            'market_divergence_20d': '%',
            'market_divergence_percentile': '%',
            'market_divergence_percentile_days': '天',
        }
        
        display_order = [
            'market_amount_20d_avg',
            '沪深300_crowding', '中证500_crowding', '中证1000_crowding',
            'market_divergence_daily', 'market_divergence_20d',
            'market_divergence_percentile', 'market_divergence_percentile_days',
        ]
        
        row = 4
        thin_border = Border(
            left=Side(style='thin'), right=Side(style='thin'),
            top=Side(style='thin'), bottom=Side(style='thin')
        )
        
        # 先按定义顺序输出
        for name in display_order:
            value = indicators.get(name)
            if value is not None:
                ws.cell(row=row, column=1, value=name)
                ws.cell(row=row, column=2, value=value)
                ws.cell(row=row, column=3, value=indicator_units.get(name, ''))
                ws.cell(row=row, column=4, value=indicator_descriptions.get(name, ''))
                
                for col in range(1, 5):
                    cell = ws.cell(row=row, column=col)
                    cell.border = thin_border
                    cell.alignment = Alignment(horizontal='left' if col in [1, 4] else 'center', vertical='center')
                
                row += 1
        
        # 输出其他未在 display_order 中的指标
        for name, value in indicators.items():
            if name not in display_order and value is not None:
                ws.cell(row=row, column=1, value=name)
                ws.cell(row=row, column=2, value=value)
                ws.cell(row=row, column=3, value=indicator_units.get(name, ''))
                ws.cell(row=row, column=4, value=indicator_descriptions.get(name, ''))
                
                for col in range(1, 5):
                    cell = ws.cell(row=row, column=col)
                    cell.border = thin_border
                    cell.alignment = Alignment(horizontal='left' if col in [1, 4] else 'center', vertical='center')
                
                row += 1
        
        ws.column_dimensions['A'].width = 35
        ws.column_dimensions['B'].width = 15
        ws.column_dimensions['C'].width = 10
        ws.column_dimensions['D'].width = 60
    
    def _create_cta_sheet(self, wb, indicators):
        """创建CTA策略指标sheet - 仅商品期货 - 优化展示格式"""
        ws = wb.create_sheet("CTA策略指标")
        
        ws['A1'] = 'CTA策略 - 商品期货指标详情'
        ws['A1'].font = Font(bold=True, size=14, color='1F4E78')
        ws.merge_cells('A1:E1')
        ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
        
        # 表头
        headers = ['品种', '指标名称', '当前值', '历史分位数', '单位']
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=3, column=col, value=header)
            cell.font = Font(bold=True, color='FFFFFF')
            cell.fill = PatternFill(start_color='ED7D31', end_color='ED7D31', fill_type='solid')
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = Border(
                left=Side(style='thin'), right=Side(style='thin'),
                top=Side(style='thin'), bottom=Side(style='thin')
            )
        
        row = 4
        
        # ===== 第一部分：波动率指标 =====
        ws.cell(row=row, column=1, value='【波动率指标】')
        ws.cell(row=row, column=1).font = Font(bold=True, size=11, color='1F4E78')
        ws.merge_cells(f'A{row}:E{row}')
        row += 1
        
        commodity_symbols = [
            ('RB', '螺纹钢'), ('CU', '铜'), ('SC', '原油'), 
            ('M', '豆粕'), ('I', '铁矿石'), ('AU', '黄金')
        ]
        
        for symbol, name in commodity_symbols:
            vol = indicators.get(f'{symbol}_volatility_20d')
            vol_pct = indicators.get(f'{symbol}_volatility_percentile')
            vol_days = indicators.get(f'{symbol}_volatility_percentile_days', 0)
            
            if vol is not None:
                ws.cell(row=row, column=1, value=name)
                ws.cell(row=row, column=2, value='20日年化波动率')
                ws.cell(row=row, column=3, value=f"{vol}%")
                if vol_pct is not None:
                    pct_display = f"{vol_pct}% (基于{vol_days}天)" if vol_days > 0 else f"{vol_pct}%"
                    ws.cell(row=row, column=4, value=pct_display)
                else:
                    ws.cell(row=row, column=4, value='数据不足')
                ws.cell(row=row, column=5, value='%')
                
                for col in range(1, 6):
                    cell = ws.cell(row=row, column=col)
                    cell.border = Border(
                        left=Side(style='thin'), right=Side(style='thin'),
                        top=Side(style='thin'), bottom=Side(style='thin')
                    )
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                row += 1
        
        # 平均波动率
        avg_vol = indicators.get('commodity_avg_volatility')
        if avg_vol is not None:
            ws.cell(row=row, column=1, value='多品种平均')
            ws.cell(row=row, column=2, value='平均波动率')
            ws.cell(row=row, column=3, value=f"{avg_vol}%")
            ws.cell(row=row, column=4, value='-')
            ws.cell(row=row, column=5, value='%')
            for col in range(1, 6):
                cell = ws.cell(row=row, column=col)
                cell.border = Border(left=Side(style='thin'), right=Side(style='thin'),
                                    top=Side(style='thin'), bottom=Side(style='thin'))
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.font = Font(bold=True)
            row += 1
        
        row += 1  # 空行
        
        # ===== 第二部分：成交量指标 =====
        ws.cell(row=row, column=1, value='【成交量指标】')
        ws.cell(row=row, column=1).font = Font(bold=True, size=11, color='1F4E78')
        ws.merge_cells(f'A{row}:E{row}')
        row += 1
        
        for symbol, name in commodity_symbols:
            vol = indicators.get(f'{symbol}_avg_volume')
            vol_pct = indicators.get(f'{symbol}_volume_percentile')
            vol_days = indicators.get(f'{symbol}_volume_percentile_days', 0)
            
            if vol is not None:
                ws.cell(row=row, column=1, value=name)
                ws.cell(row=row, column=2, value='近20日平均成交量')
                ws.cell(row=row, column=3, value=f"{vol:,.0f}")
                if vol_pct is not None:
                    pct_display = f"{vol_pct}% (基于{vol_days}天)" if vol_days > 0 else f"{vol_pct}%"
                    ws.cell(row=row, column=4, value=pct_display)
                else:
                    ws.cell(row=row, column=4, value='数据不足')
                ws.cell(row=row, column=5, value='手')
                
                for col in range(1, 6):
                    cell = ws.cell(row=row, column=col)
                    cell.border = Border(
                        left=Side(style='thin'), right=Side(style='thin'),
                        top=Side(style='thin'), bottom=Side(style='thin')
                    )
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                row += 1
        
        # 总成交量
        total_vol = indicators.get('commodity_total_volume')
        if total_vol is not None:
            ws.cell(row=row, column=1, value='多品种合计')
            ws.cell(row=row, column=2, value='总成交量')
            ws.cell(row=row, column=3, value=f"{total_vol:,.0f}")
            ws.cell(row=row, column=4, value='-')
            ws.cell(row=row, column=5, value='手')
            for col in range(1, 6):
                cell = ws.cell(row=row, column=col)
                cell.border = Border(left=Side(style='thin'), right=Side(style='thin'),
                                    top=Side(style='thin'), bottom=Side(style='thin'))
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.font = Font(bold=True)
            row += 1
        
        row += 1  # 空行
        
        # ===== 第三部分：持仓量指标 =====
        ws.cell(row=row, column=1, value='【持仓量指标】')
        ws.cell(row=row, column=1).font = Font(bold=True, size=11, color='1F4E78')
        ws.merge_cells(f'A{row}:E{row}')
        row += 1
        
        for symbol, name in commodity_symbols:
            pos = indicators.get(f'{symbol}_positions')
            pos_pct = indicators.get(f'{symbol}_positions_percentile')
            pos_days = indicators.get(f'{symbol}_positions_percentile_days', 0)
            
            if pos is not None:
                ws.cell(row=row, column=1, value=name)
                ws.cell(row=row, column=2, value='最新持仓量')
                ws.cell(row=row, column=3, value=f"{pos:,.0f}")
                if pos_pct is not None:
                    pct_display = f"{pos_pct}% (基于{pos_days}天)" if pos_days > 0 else f"{pos_pct}%"
                    ws.cell(row=row, column=4, value=pct_display)
                else:
                    ws.cell(row=row, column=4, value='数据不足')
                ws.cell(row=row, column=5, value='手')
                
                for col in range(1, 6):
                    cell = ws.cell(row=row, column=col)
                    cell.border = Border(
                        left=Side(style='thin'), right=Side(style='thin'),
                        top=Side(style='thin'), bottom=Side(style='thin')
                    )
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                row += 1
        
        # 总持仓量
        total_pos = indicators.get('commodity_total_positions')
        if total_pos is not None:
            ws.cell(row=row, column=1, value='多品种合计')
            ws.cell(row=row, column=2, value='总持仓量')
            ws.cell(row=row, column=3, value=f"{total_pos:,.0f}")
            ws.cell(row=row, column=4, value='-')
            ws.cell(row=row, column=5, value='手')
            for col in range(1, 6):
                cell = ws.cell(row=row, column=col)
                cell.border = Border(left=Side(style='thin'), right=Side(style='thin'),
                                    top=Side(style='thin'), bottom=Side(style='thin'))
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.font = Font(bold=True)
            row += 1
        
        # 设置列宽
        ws.column_dimensions['A'].width = 12
        ws.column_dimensions['B'].width = 18
        ws.column_dimensions['C'].width = 18
        ws.column_dimensions['D'].width = 22
        ws.column_dimensions['E'].width = 8
    
    def _create_arbitrage_sheet(self, wb, indicators):
        """创建套利策略指标sheet - 优化展示格式"""
        ws = wb.create_sheet("套利策略指标")
        
        ws['A1'] = '套利策略 - 指标详情（ETF套利/股指高频套利/期权套利/基差）'
        ws['A1'].font = Font(bold=True, size=14, color='1F4E78')
        ws.merge_cells('A1:E1')
        ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
        
        # 表头
        headers = ['标的', '指标名称', '当前值', '历史分位数', '单位']
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=3, column=col, value=header)
            cell.font = Font(bold=True, color='FFFFFF')
            cell.fill = PatternFill(start_color='A5A5A5', end_color='A5A5A5', fill_type='solid')
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = Border(
                left=Side(style='thin'), right=Side(style='thin'),
                top=Side(style='thin'), bottom=Side(style='thin')
            )
        
        row = 4
        
        # ===== 第一部分：ETF套利指标 =====
        ws.cell(row=row, column=1, value='【ETF套利指标】')
        ws.cell(row=row, column=1).font = Font(bold=True, size=11, color='1F4E78')
        ws.merge_cells(f'A{row}:E{row}')
        row += 1
        
        etf_list = [
            ('300ETF', '沪深300ETF'), ('500ETF', '中证500ETF'), ('1000ETF', '中证1000ETF'),
            ('50ETF', '上证50ETF'), ('创业板ETF', '创业板ETF'), ('科创50ETF', '科创50ETF')
        ]
        
        for etf_code, etf_name in etf_list:
            # 成交额
            amount = indicators.get(f'{etf_code}_avg_amount_20d')
            amount_pct = indicators.get(f'{etf_code}_amount_percentile')
            amount_days = indicators.get(f'{etf_code}_amount_percentile_days', 0)
            
            if amount is not None:
                ws.cell(row=row, column=1, value=etf_name)
                ws.cell(row=row, column=2, value='近20日平均成交额')
                ws.cell(row=row, column=3, value=f"{amount:.2f}")
                if amount_pct is not None:
                    pct_display = f"{amount_pct}% (基于{amount_days}天)" if amount_days > 0 else f"{amount_pct}%"
                    ws.cell(row=row, column=4, value=pct_display)
                else:
                    ws.cell(row=row, column=4, value='数据不足')
                ws.cell(row=row, column=5, value='亿元')
                for col in range(1, 6):
                    cell = ws.cell(row=row, column=col)
                    cell.border = Border(left=Side(style='thin'), right=Side(style='thin'),
                                        top=Side(style='thin'), bottom=Side(style='thin'))
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                row += 1
            
            # 波动率
            vol = indicators.get(f'{etf_code}_volatility_20d')
            vol_pct = indicators.get(f'{etf_code}_volatility_percentile')
            vol_days = indicators.get(f'{etf_code}_volatility_percentile_days', 0)
            
            if vol is not None:
                ws.cell(row=row, column=1, value=etf_name)
                ws.cell(row=row, column=2, value='20日年化波动率')
                ws.cell(row=row, column=3, value=f"{vol}%")
                if vol_pct is not None:
                    pct_display = f"{vol_pct}% (基于{vol_days}天)" if vol_days > 0 else f"{vol_pct}%"
                    ws.cell(row=row, column=4, value=pct_display)
                else:
                    ws.cell(row=row, column=4, value='数据不足')
                ws.cell(row=row, column=5, value='%')
                for col in range(1, 6):
                    cell = ws.cell(row=row, column=col)
                    cell.border = Border(left=Side(style='thin'), right=Side(style='thin'),
                                        top=Side(style='thin'), bottom=Side(style='thin'))
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                row += 1
            
            # 折溢价率
            premium = indicators.get(f'{etf_code}_premium')
            premium_pct = indicators.get(f'{etf_code}_premium_percentile')
            premium_days = indicators.get(f'{etf_code}_premium_percentile_days', 0)
            
            if premium is not None:
                ws.cell(row=row, column=1, value=etf_name)
                ws.cell(row=row, column=2, value='折溢价率')
                ws.cell(row=row, column=3, value=f"{premium:.4f}%")
                if premium_pct is not None:
                    pct_display = f"{premium_pct}% (基于{premium_days}天)" if premium_days > 0 else f"{premium_pct}%"
                    ws.cell(row=row, column=4, value=pct_display)
                else:
                    ws.cell(row=row, column=4, value='数据不足')
                ws.cell(row=row, column=5, value='%')
                # 折溢价为负（折价）时标绿色，为正（溢价）时标红色
                fill_color = 'C6EFCE' if premium < 0 else 'FFC7CE'
                for col in range(1, 6):
                    cell = ws.cell(row=row, column=col)
                    cell.border = Border(left=Side(style='thin'), right=Side(style='thin'),
                                        top=Side(style='thin'), bottom=Side(style='thin'))
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                    cell.fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type='solid')
                row += 1
        
        row += 1  # 空行
        
        # ===== 第二部分：股指高频套利指标 =====
        ws.cell(row=row, column=1, value='【股指高频套利指标】')
        ws.cell(row=row, column=1).font = Font(bold=True, size=11, color='1F4E78')
        ws.merge_cells(f'A{row}:E{row}')
        row += 1
        
        futures_list = [
            ('IF', '沪深300股指'), ('IC', '中证500股指'), ('IM', '中证1000股指')
        ]
        
        for fut_code, fut_name in futures_list:
            # 波动率
            vol = indicators.get(f'{fut_code}_volatility_20d')
            vol_pct = indicators.get(f'{fut_code}_volatility_percentile')
            vol_days = indicators.get(f'{fut_code}_volatility_percentile_days', 0)
            
            if vol is not None:
                ws.cell(row=row, column=1, value=fut_name)
                ws.cell(row=row, column=2, value='20日年化波动率')
                ws.cell(row=row, column=3, value=f"{vol}%")
                if vol_pct is not None:
                    pct_display = f"{vol_pct}% (基于{vol_days}天)" if vol_days > 0 else f"{vol_pct}%"
                    ws.cell(row=row, column=4, value=pct_display)
                else:
                    ws.cell(row=row, column=4, value='数据不足')
                ws.cell(row=row, column=5, value='%')
                for col in range(1, 6):
                    cell = ws.cell(row=row, column=col)
                    cell.border = Border(left=Side(style='thin'), right=Side(style='thin'),
                                        top=Side(style='thin'), bottom=Side(style='thin'))
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                row += 1
            
            # 成交额分位数
            amt_pct = indicators.get(f'{fut_code}_amount_percentile')
            amt_days = indicators.get(f'{fut_code}_amount_percentile_days', 0)
            
            if amt_pct is not None:
                ws.cell(row=row, column=1, value=fut_name)
                ws.cell(row=row, column=2, value='成交额历史分位数')
                ws.cell(row=row, column=3, value='-')
                pct_display = f"{amt_pct}% (基于{amt_days}天)" if amt_days > 0 else f"{amt_pct}%"
                ws.cell(row=row, column=4, value=pct_display)
                ws.cell(row=row, column=5, value='%')
                for col in range(1, 6):
                    cell = ws.cell(row=row, column=col)
                    cell.border = Border(left=Side(style='thin'), right=Side(style='thin'),
                                        top=Side(style='thin'), bottom=Side(style='thin'))
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                row += 1
        
        row += 1  # 空行
        
        # ===== 第三部分：期权套利指标 =====
        ws.cell(row=row, column=1, value='【期权套利指标 - 隐含波动率IV】')
        ws.cell(row=row, column=1).font = Font(bold=True, size=11, color='1F4E78')
        ws.merge_cells(f'A{row}:E{row}')
        row += 1
        
        option_list = [
            ('50ETF', '50ETF期权'), ('300ETF', '300ETF期权'), ('500ETF', '500ETF期权'),
            ('科创50ETF', '科创50ETF期权'), ('创业板ETF', '创业板ETF期权'),
            ('300股指', '300股指期权'), ('1000股指', '1000股指期权'), ('50股指', '50股指期权')
        ]
        
        for opt_code, opt_name in option_list:
            iv = indicators.get(f'{opt_code}_IV')
            iv_pct = indicators.get(f'{opt_code}_IV_percentile')
            iv_days = indicators.get(f'{opt_code}_IV_percentile_days', 0)
            
            if iv is not None:
                ws.cell(row=row, column=1, value=opt_name)
                ws.cell(row=row, column=2, value='隐含波动率')
                ws.cell(row=row, column=3, value=f"{iv}%")
                if iv_pct is not None:
                    pct_display = f"{iv_pct}% (基于{iv_days}天)" if iv_days > 0 else f"{iv_pct}%"
                    ws.cell(row=row, column=4, value=pct_display)
                else:
                    ws.cell(row=row, column=4, value='数据不足')
                ws.cell(row=row, column=5, value='%')
                for col in range(1, 6):
                    cell = ws.cell(row=row, column=col)
                    cell.border = Border(left=Side(style='thin'), right=Side(style='thin'),
                                        top=Side(style='thin'), bottom=Side(style='thin'))
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                row += 1
        
        row += 1  # 空行
        
        # ===== 第四部分：股指期货基差 =====
        ws.cell(row=row, column=1, value='【股指期货基差】')
        ws.cell(row=row, column=1).font = Font(bold=True, size=11, color='1F4E78')
        ws.merge_cells(f'A{row}:E{row}')
        row += 1
        
        basis_list = [
            ('IF', '沪深300股指'), ('IC', '中证500股指'), ('IM', '中证1000股指')
        ]
        
        for fut_code, fut_name in basis_list:
            basis = indicators.get(f'{fut_code}_basis')
            basis_annual = indicators.get(f'{fut_code}_basis_annual')
            
            if basis is not None:
                ws.cell(row=row, column=1, value=fut_name)
                ws.cell(row=row, column=2, value='基差')
                ws.cell(row=row, column=3, value=f"{basis:.4f}%")
                ws.cell(row=row, column=4, value='-')
                ws.cell(row=row, column=5, value='%')
                # 基差为负（贴水）时标绿色，为正（升水）时标红色
                fill_color = 'C6EFCE' if basis < 0 else 'FFC7CE'
                for col in range(1, 6):
                    cell = ws.cell(row=row, column=col)
                    cell.border = Border(left=Side(style='thin'), right=Side(style='thin'),
                                        top=Side(style='thin'), bottom=Side(style='thin'))
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                    cell.fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type='solid')
                row += 1
            
            if basis_annual is not None:
                ws.cell(row=row, column=1, value=fut_name)
                ws.cell(row=row, column=2, value='年化基差')
                ws.cell(row=row, column=3, value=f"{basis_annual:.2f}%")
                ws.cell(row=row, column=4, value='-')
                ws.cell(row=row, column=5, value='%')
                fill_color = 'C6EFCE' if basis_annual < 0 else 'FFC7CE'
                for col in range(1, 6):
                    cell = ws.cell(row=row, column=col)
                    cell.border = Border(left=Side(style='thin'), right=Side(style='thin'),
                                        top=Side(style='thin'), bottom=Side(style='thin'))
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                    cell.fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type='solid')
                row += 1
        
        # 设置列宽
        ws.column_dimensions['A'].width = 14
        ws.column_dimensions['B'].width = 18
        ws.column_dimensions['C'].width = 16
        ws.column_dimensions['D'].width = 22
        ws.column_dimensions['E'].width = 8
    
    def _create_neutral_sheet(self, wb, indicators):
        """创建市场中性策略指标sheet - 优化展示格式"""
        ws = wb.create_sheet("市场中性指标")
        
        ws['A1'] = '市场中性策略 - 指标详情'
        ws['A1'].font = Font(bold=True, size=14, color='1F4E78')
        ws.merge_cells('A1:D1')
        ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
        
        # 表头
        headers = ['指标名称', '当前值', '单位', '说明']
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=3, column=col, value=header)
            cell.font = Font(bold=True, color='FFFFFF')
            cell.fill = PatternFill(start_color='7030A0', end_color='7030A0', fill_type='solid')
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = Border(
                left=Side(style='thin'), right=Side(style='thin'),
                top=Side(style='thin'), bottom=Side(style='thin')
            )
        
        row = 4
        
        # ===== 第一部分：市场活跃度 =====
        ws.cell(row=row, column=1, value='【市场活跃度】')
        ws.cell(row=row, column=1).font = Font(bold=True, size=11, color='1F4E78')
        ws.merge_cells(f'A{row}:D{row}')
        row += 1
        
        market_amount = indicators.get('market_amount_20d_avg')
        if market_amount is not None:
            ws.cell(row=row, column=1, value='近20日市场平均成交额')
            ws.cell(row=row, column=2, value=f"{market_amount:.2f}")
            ws.cell(row=row, column=3, value='亿元')
            ws.cell(row=row, column=4, value='反映市场整体活跃度')
            for col in range(1, 5):
                cell = ws.cell(row=row, column=col)
                cell.border = Border(left=Side(style='thin'), right=Side(style='thin'),
                                    top=Side(style='thin'), bottom=Side(style='thin'))
                cell.alignment = Alignment(horizontal='left' if col in [1, 4] else 'center', vertical='center')
            row += 1
        
        row += 1  # 空行
        
        # ===== 第二部分：指数拥挤度 =====
        ws.cell(row=row, column=1, value='【指数拥挤度 - 成交额占全市场比例】')
        ws.cell(row=row, column=1).font = Font(bold=True, size=11, color='1F4E78')
        ws.merge_cells(f'A{row}:D{row}')
        row += 1
        
        index_list = [
            ('沪深300', '大盘'), ('中证500', '中盘'), ('中证1000', '小盘')
        ]
        
        for idx_name, idx_type in index_list:
            crowding = indicators.get(f'{idx_name}_crowding')
            if crowding is not None:
                ws.cell(row=row, column=1, value=f'{idx_name}（{idx_type}）')
                ws.cell(row=row, column=2, value=f"{crowding:.2f}%")
                ws.cell(row=row, column=3, value='%')
                ws.cell(row=row, column=4, value=f'{idx_name}成交额/全市场总成交额')
                for col in range(1, 5):
                    cell = ws.cell(row=row, column=col)
                    cell.border = Border(left=Side(style='thin'), right=Side(style='thin'),
                                        top=Side(style='thin'), bottom=Side(style='thin'))
                    cell.alignment = Alignment(horizontal='left' if col in [1, 4] else 'center', vertical='center')
                row += 1
        
        row += 1  # 空行
        
        # ===== 第三部分：股指期货基差 =====
        ws.cell(row=row, column=1, value='【股指期货基差 - 对冲成本/收益】')
        ws.cell(row=row, column=1).font = Font(bold=True, size=11, color='1F4E78')
        ws.merge_cells(f'A{row}:D{row}')
        row += 1
        
        # 说明行
        ws.cell(row=row, column=1, value='说明')
        ws.cell(row=row, column=1).font = Font(italic=True, color='666666')
        ws.cell(row=row, column=2, value='基差=(期货-现货)/现货；负为贴水（对冲成本），正为升水（对冲收益）')
        ws.merge_cells(f'B{row}:D{row}')
        ws.cell(row=row, column=2).font = Font(italic=True, color='666666')
        row += 1
        
        basis_list = [
            ('IF', '沪深300股指'), ('IC', '中证500股指'), ('IM', '中证1000股指')
        ]
        
        for fut_code, fut_name in basis_list:
            basis = indicators.get(f'{fut_code}_basis')
            basis_annual = indicators.get(f'{fut_code}_basis_annual')
            
            if basis is not None:
                ws.cell(row=row, column=1, value=f'{fut_name}基差')
                ws.cell(row=row, column=2, value=f"{basis:.4f}%")
                ws.cell(row=row, column=3, value='%')
                ws.cell(row=row, column=4, value='当前基差水平')
                fill_color = 'C6EFCE' if basis < 0 else 'FFC7CE'
                for col in range(1, 5):
                    cell = ws.cell(row=row, column=col)
                    cell.border = Border(left=Side(style='thin'), right=Side(style='thin'),
                                        top=Side(style='thin'), bottom=Side(style='thin'))
                    cell.alignment = Alignment(horizontal='left' if col in [1, 4] else 'center', vertical='center')
                    cell.fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type='solid')
                row += 1
            
            if basis_annual is not None:
                ws.cell(row=row, column=1, value=f'{fut_name}年化基差')
                ws.cell(row=row, column=2, value=f"{basis_annual:.2f}%")
                ws.cell(row=row, column=3, value='%')
                desc = '对冲有收益' if basis_annual > 0 else '对冲有成本'
                ws.cell(row=row, column=4, value=desc)
                fill_color = 'C6EFCE' if basis_annual > 0 else 'FFC7CE'
                for col in range(1, 5):
                    cell = ws.cell(row=row, column=col)
                    cell.border = Border(left=Side(style='thin'), right=Side(style='thin'),
                                        top=Side(style='thin'), bottom=Side(style='thin'))
                    cell.alignment = Alignment(horizontal='left' if col in [1, 4] else 'center', vertical='center')
                    cell.fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type='solid')
                row += 1
        
        row += 1  # 空行
        
        # ===== 第四部分：市场形态指标 =====
        ws.cell(row=row, column=1, value='【市场形态指标 - 哑铃型/纺锤型】')
        ws.cell(row=row, column=1).font = Font(bold=True, size=11, color='1F4E78')
        ws.merge_cells(f'A{row}:D{row}')
        row += 1
        
        # 说明行
        ws.cell(row=row, column=1, value='说明')
        ws.cell(row=row, column=1).font = Font(italic=True, color='666666')
        ws.cell(row=row, column=2, value='哑铃指数=国证2000+沪深300-中证500；>MA5为哑铃型，<MA5为纺锤型')
        ws.merge_cells(f'B{row}:D{row}')
        ws.cell(row=row, column=2).font = Font(italic=True, color='666666')
        row += 1
        
        # 哑铃指数
        dumbbell_index = indicators.get('dumbbell_index')
        if dumbbell_index is not None:
            ws.cell(row=row, column=1, value='哑铃指数')
            ws.cell(row=row, column=2, value=f"{dumbbell_index:.4f}%")
            ws.cell(row=row, column=3, value='%')
            ws.cell(row=row, column=4, value='国证2000+沪深300-中证500')
            for col in range(1, 5):
                cell = ws.cell(row=row, column=col)
                cell.border = Border(left=Side(style='thin'), right=Side(style='thin'),
                                    top=Side(style='thin'), bottom=Side(style='thin'))
                cell.alignment = Alignment(horizontal='left' if col in [1, 4] else 'center', vertical='center')
            row += 1
        
        # 5日均线
        dumbbell_ma5 = indicators.get('dumbbell_ma5')
        if dumbbell_ma5 is not None:
            ws.cell(row=row, column=1, value='哑铃指数5日均线')
            ws.cell(row=row, column=2, value=f"{dumbbell_ma5:.4f}%")
            ws.cell(row=row, column=3, value='%')
            ws.cell(row=row, column=4, value='过去5个交易日均值')
            for col in range(1, 5):
                cell = ws.cell(row=row, column=col)
                cell.border = Border(left=Side(style='thin'), right=Side(style='thin'),
                                    top=Side(style='thin'), bottom=Side(style='thin'))
                cell.alignment = Alignment(horizontal='left' if col in [1, 4] else 'center', vertical='center')
            row += 1
        
        # 市场形态
        pattern = indicators.get('market_pattern')
        if pattern is not None:
            ws.cell(row=row, column=1, value='市场形态')
            ws.cell(row=row, column=2, value=pattern)
            ws.cell(row=row, column=3, value='-')
            if pattern == '哑铃型':
                desc = '两头强于中间，风格分化大，对冲难度↑'
                fill_color = 'FFC7CE'  # 红色（不利）
            elif pattern == '纺锤型':
                desc = '中间强于两头，风格集中，对冲难度↓'
                fill_color = 'C6EFCE'  # 绿色（有利）
            else:
                desc = '风格不明显，市场环境稳定'
                fill_color = None
            ws.cell(row=row, column=4, value=desc)
            for col in range(1, 5):
                cell = ws.cell(row=row, column=col)
                cell.border = Border(left=Side(style='thin'), right=Side(style='thin'),
                                    top=Side(style='thin'), bottom=Side(style='thin'))
                cell.alignment = Alignment(horizontal='left' if col in [1, 4] else 'center', vertical='center')
                if fill_color:
                    cell.fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type='solid')
            row += 1
        
        # 设置列宽
        ws.column_dimensions['A'].width = 22
        ws.column_dimensions['B'].width = 16
        ws.column_dimensions['C'].width = 8
        ws.column_dimensions['D'].width = 40
    
    def _create_score_details_sheet(self, wb, all_scores):
        """创建评分详情sheet"""
        ws = wb.create_sheet("评分详情")
        
        ws['A1'] = '策略评分详细说明'
        ws['A1'].font = Font(bold=True, size=14, color='1F4E78')
        ws.merge_cells('A1:C1')
        ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
        
        headers = ['策略类型', '评分', '详细说明']
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=3, column=col, value=header)
            cell.font = Font(bold=True, color='FFFFFF')
            cell.fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = Border(
                left=Side(style='thin'), right=Side(style='thin'),
                top=Side(style='thin'), bottom=Side(style='thin')
            )
        
        row = 4
        for strategy_name, result in all_scores.items():
            score = result['score']
            details = result.get('details', [])
            
            ws.cell(row=row, column=1, value=strategy_name)
            ws.cell(row=row, column=2, value=score)
            
            detail_text = '\n'.join([f'• {d}' for d in details])
            ws.cell(row=row, column=3, value=detail_text)
            
            if score >= 80:
                fill_color = 'C6EFCE'
            elif score >= 65:
                fill_color = 'E2EFDA'
            elif score >= 45:
                fill_color = 'FFEB9C'
            elif score >= 30:
                fill_color = 'FFC7CE'
            else:
                fill_color = 'FF6B6B'
            
            for col in range(1, 4):
                cell = ws.cell(row=row, column=col)
                cell.fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type='solid')
                cell.border = Border(
                    left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin')
                )
                cell.alignment = Alignment(horizontal='center' if col <= 2 else 'left', vertical='top', wrap_text=True)
            
            ws.row_dimensions[row].height = max(30, len(details) * 15)
            
            row += 1
        
        ws.column_dimensions['A'].width = 15
        ws.column_dimensions['B'].width = 10
        ws.column_dimensions['C'].width = 60


# ============================================
# 主程序 - 聚宽研究环境入口
# ============================================

def collect_all_indicators():
    """收集所有策略的指标"""
    print("=" * 60)
    print("正在收集各策略指标...")
    print("=" * 60)
    
    calculator = IndicatorCalculator()
    all_indicators = {}
    
    # 1. 主观多头指标
    print("\n[1/7] 收集主观多头指标...")
    subjective = {}
    subjective.update(calculator.get_sentiment_indicators())
    subjective.update(calculator.get_valuation_indicators())
    all_indicators['subjective'] = subjective
    print(f"  - 市场情绪指标: {len([k for k in subjective.keys() if 'amount' in k])}个")
    print(f"  - 估值指标: {len([k for k in subjective.keys() if 'percentile' in k])}个")
    
    # 2. 量化多头指标
    print("\n[2/7] 收集量化多头指标...")
    quant = calculator.get_quant_indicators()
    all_indicators['quant'] = quant
    print(f"  - 拥挤度指标: {len([k for k in quant.keys() if 'crowding' in k])}个")
    print(f"  - 分化度相关指标: {len([k for k in quant.keys() if 'divergence' in k])}个")
    if quant.get('market_divergence_daily') is not None:
        print(f"  - 当日市值加权-等权分化度: {quant.get('market_divergence_daily')}")
    
    # 3. CTA指标（仅商品期货）
    print("\n[3/7] 收集CTA策略指标（商品期货）...")
    cta = calculator.get_cta_indicators()
    all_indicators['cta'] = cta
    print(f"  - 商品期货波动率指标: {len([k for k in cta.keys() if 'volatility' in k])}个")
    print(f"  - 商品期货成交量指标: {len([k for k in cta.keys() if 'volume' in k])}个")
    print(f"  - 商品期货持仓量指标: {len([k for k in cta.keys() if 'positions' in k])}个")
    
    # 4. ETF套利指标
    print("\n[4/7] 收集ETF套利指标...")
    arbitrage = calculator.get_arbitrage_indicators()
    all_indicators['arbitrage'] = arbitrage
    print(f"  - ETF套利指标: {len([k for k in arbitrage.keys() if 'ETF' in k])}个")
    
    # 5. 股指套利指标
    print("\n[5/7] 收集股指套利指标...")
    print(f"  - 股指波动率指标: {len([k for k in arbitrage.keys() if 'volatility' in k and 'ETF' not in k and 'IV' not in k])}个")
    print(f"  - 基差指标: {len([k for k in arbitrage.keys() if 'basis' in k])}个")
    
    # 6. 期权套利指标
    print("\n[6/7] 收集期权套利指标...")
    print(f"  - 期权IV指标: {len([k for k in arbitrage.keys() if 'IV' in k])}个")
    
    # 7. 市场中性指标
    print("\n[7/7] 收集市场中性指标...")
    neutral = calculator.get_market_neutral_indicators()
    all_indicators['neutral'] = neutral
    print(f"  - 基差指标: {len([k for k in neutral.keys() if 'basis' in k])}个")
    print(f"  - 拥挤度指标: {len([k for k in neutral.keys() if 'crowding' in k])}个")
    
    return all_indicators


def calculate_all_scores():
    """计算所有策略评分"""
    print("\n" + "=" * 60)
    print("正在计算策略评分...")
    print("=" * 60)
    
    scorer = StrategyScorer()
    all_scores = scorer.get_all_scores()
    
    for strategy_name, result in all_scores.items():
        print(f"\n{strategy_name}: {result['score']}分 - {result['level']}")
        for detail in result['details']:
            print(f"  • {detail}")
    
    return all_scores


def generate_report(all_scores, all_indicators):
    """生成控制台报告"""
    print("\n" + "=" * 60)
    print("策略评分总览")
    print("=" * 60)
    print(f"{'策略类型':<12} {'评分':<8} {'适配环境':<12} {'趋势':<8}")
    print("-" * 45)
    for strategy_name, result in all_scores.items():
        print(f"{strategy_name:<12} {result['score']:<8} {result['level']:<12} {result['trend']:<8}")


# ============================================
# 聚宽研究环境入口函数
# ============================================

def run_strategy_monitor(output_excel=True, excel_path=None):
    """
    策略环境监测主函数
    :param output_excel: 是否输出Excel报告
    :param excel_path: Excel文件保存路径，默认自动生成
    :return: (all_scores, all_indicators, excel_path)
    """
    print("\n" + "=" * 60)
    print("策略环境监测系统")
    print(f"运行时间: {dt.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 60)
    
    # 1. 收集所有指标
    all_indicators = collect_all_indicators()
    
    # 2. 计算所有策略评分
    all_scores = calculate_all_scores()
    
    # 3. 生成控制台报告
    generate_report(all_scores, all_indicators)
    
    # 4. 生成Excel报告
    excel_file_path = None
    if output_excel:
        print("\n" + "=" * 60)
        print("正在生成Excel报告...")
        print("=" * 60)
        generator = ExcelReportGenerator(output_path=excel_path)
        excel_file_path = generator.generate_excel_report(all_scores, all_indicators)
    
    print("\n" + "=" * 60)
    print("分析完成!")
    print("=" * 60)
    
    return all_scores, all_indicators, excel_file_path


# ============================================
# 运行入口
# ============================================

if __name__ == "__main__":
    # 在聚宽研究环境中直接运行
    scores, indicators, excel_path = run_strategy_monitor(output_excel=True)
