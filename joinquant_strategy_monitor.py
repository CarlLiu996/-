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


def get_weekly_observation_end_date(run_dt=None):
    """周度监测：数据截止日 = 严格早于运行日的「最近一个自然周五」对应的最近 A 股交易日（不晚于该周五）。"""
    if run_dt is None:
        run_dt = dt.datetime.now()
    cal = run_dt.date()
    target_friday = None
    for i in range(1, 21):
        c = cal - dt.timedelta(days=i)
        if c.weekday() == 4:
            target_friday = c
            break
    if target_friday is None:
        target_friday = cal - dt.timedelta(days=7)
    target_s = target_friday.strftime('%Y-%m-%d')
    try:
        start_s = (target_friday - dt.timedelta(days=45)).strftime('%Y-%m-%d')
        try:
            tds = get_trade_days(start_date=start_s, end_date=target_s)
        except TypeError:
            tds = get_trade_days(end_date=target_s, count=30)
        if tds is not None and len(tds) > 0:
            last = tds[-1]
            if hasattr(last, 'strftime'):
                return str(last)[:10]
            return pd.Timestamp(last).strftime('%Y-%m-%d')
    except Exception:
        pass
    return target_s


def _resample_ohlcv_to_weekly_friday(df, date_col='日期'):
    """日线 DataFrame → 周线（周五收盘），成交量/额取周内合计。"""
    if df is None or df.empty or date_col not in df.columns:
        return None
    d = df.sort_values(date_col).copy()
    d[date_col] = pd.to_datetime(d[date_col])
    d = d.set_index(date_col)
    agg = {}
    if '开盘' in d.columns:
        agg['开盘'] = 'first'
    if '收盘' in d.columns:
        agg['收盘'] = 'last'
    if '最高' in d.columns:
        agg['最高'] = 'max'
    if '最低' in d.columns:
        agg['最低'] = 'min'
    if '成交量' in d.columns:
        agg['成交量'] = 'sum'
    if '成交额' in d.columns:
        agg['成交额'] = 'sum'
    if '持仓量' in d.columns:
        agg['持仓量'] = 'sum'
    if not agg:
        return None
    w = d.resample('W-FRI').agg(agg).dropna(how='all')
    if w.empty:
        return None
    out = w.reset_index()
    if out.columns[0] != date_col:
        out = out.rename(columns={out.columns[0]: date_col})
    return out


def _resample_futures_daily_to_weekly(df, date_col='日期'):
    """期货主力日线（开盘价/收盘价…）→ 周线。"""
    if df is None or df.empty or date_col not in df.columns:
        return None
    rename_close = '收盘价' in df.columns
    d = df.sort_values(date_col).copy()
    d[date_col] = pd.to_datetime(d[date_col])
    d = d.set_index(date_col)
    agg = {}
    if '开盘价' in d.columns:
        agg['开盘价'] = 'first'
    if '收盘价' in d.columns:
        agg['收盘价'] = 'last'
    if '最高价' in d.columns:
        agg['最高价'] = 'max'
    if '最低价' in d.columns:
        agg['最低价'] = 'min'
    if '成交量' in d.columns:
        agg['成交量'] = 'sum'
    if '成交额' in d.columns:
        agg['成交额'] = 'sum'
    if not agg:
        return None
    w = d.resample('W-FRI').agg(agg).dropna(how='all')
    if w.empty:
        return None
    out = w.reset_index()
    if out.columns[0] != date_col:
        out = out.rename(columns={out.columns[0]: date_col})
    if rename_close and '最新价' not in out.columns and '收盘价' in out.columns:
        out['最新价'] = out['收盘价']
    return out


# ============================================
# 数据获取模块 - 使用聚宽API
# ============================================

class DataFetcher:
    """数据获取类 - 聚宽版本"""
    
    def __init__(self, observation_end_date=None):
        """observation_end_date: 'YYYY-MM-DD'，为 None 时用运行当日（兼容旧行为）。"""
        self.observation_end_date = observation_end_date
    
    def _end(self, end_date):
        if end_date is not None:
            return end_date
        if self.observation_end_date:
            return self.observation_end_date
        return dt.datetime.now().strftime('%Y-%m-%d')
    
    def get_index_daily(self, symbol, period=120, end_date=None):
        """获取指数日线数据"""
        end_date = self._end(end_date)
        
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
        end_date = self._end(end_date)
        
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
        end_date = self._end(end_date)
        
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
    
    def get_futures_main_contract(self, underlying_symbol, period=252, end_date=None):
        """获取期货主力合约数据"""
        try:
            ed = self._end(end_date)
            try:
                dominant = get_dominant_future(underlying_symbol, ed)
            except TypeError:
                dominant = get_dominant_future(underlying_symbol)
            if dominant:
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
    
    def get_index_weekly(self, symbol, approx_weeks=140, end_date=None):
        """指数周线（周五），由日线重采样。"""
        daily_count = max(300, approx_weeks * 7 + 80)
        df = self.get_index_daily(symbol, period=daily_count, end_date=end_date)
        return _resample_ohlcv_to_weekly_friday(df)

    def get_etf_weekly(self, symbol, approx_weeks=140, end_date=None):
        daily_count = max(300, approx_weeks * 7 + 80)
        df = self.get_etf_daily(symbol, period=daily_count, end_date=end_date)
        return _resample_ohlcv_to_weekly_friday(df)

    def get_futures_main_weekly(self, underlying_symbol, approx_weeks=140, end_date=None):
        daily_count = max(400, approx_weeks * 7 + 80)
        df = self.get_futures_main_contract(underlying_symbol, period=daily_count, end_date=end_date)
        return _resample_futures_daily_to_weekly(df)

    def get_futures_info_data(self, underlying_symbol):
        """获取期货合约信息和持仓量"""
        try:
            ed = self._end(None)
            try:
                dominant = get_dominant_future(underlying_symbol, ed)
            except TypeError:
                dominant = get_dominant_future(underlying_symbol)
            if dominant:
                end_dt = dt.datetime.strptime(ed, '%Y-%m-%d')
                df = get_extras('futures_positions', [dominant], 
                               start_date=(end_dt - dt.timedelta(days=252)).strftime('%Y-%m-%d'),
                               end_date=ed)
                return df
            return None
        except Exception as e:
            print(f"获取期货持仓量失败 {underlying_symbol}: {e}")
            return None

    def get_commodity_equal_weight_main_daily(self, period=600, end_date=None):
        """非中金所商品期货：各品种主力合约收盘价归一化后等权合成指数，成交量/持仓为各品种合计。"""

        def _underlying_from_code(sec_code):
            base = str(sec_code).split('.')[0]
            letters = []
            for ch in base:
                if ch.isalpha():
                    letters.append(ch)
                else:
                    break
            return ''.join(letters).upper() if letters else None

        ed = self._end(end_date)
        try:
            try:
                all_f = get_all_securities(types=['futures'], date=ed)
            except TypeError:
                all_f = get_all_securities(['futures'])
        except Exception as e:
            print(f"获取期货列表失败: {e}")
            return None
        if all_f is None or all_f.empty:
            return None

        under_set = set()
        for code in all_f.index:
            sc = str(code)
            if sc.endswith('.CCFX'):
                continue
            u = _underlying_from_code(sc)
            if u:
                under_set.add(u)
        underlyings = sorted(under_set)

        closes, vols, ois = {}, {}, {}
        for u in underlyings:
            try:
                try:
                    main = get_dominant_future(u, ed)
                except TypeError:
                    main = get_dominant_future(u)
                if not main:
                    continue
                try:
                    df = get_price(
                        main, count=period, end_date=ed, frequency='daily',
                        fields=['close', 'volume', 'open_interest'])
                except Exception:
                    df = get_price(
                        main, count=period, end_date=ed, frequency='daily',
                        fields=['close', 'volume'])
                if df is None or df.empty:
                    continue
                df = df.reset_index()
                tcol = 'index' if 'index' in df.columns else df.columns[0]
                df.rename(columns={tcol: '日期', 'close': '收盘', 'volume': '成交量'}, inplace=True)
                if 'open_interest' in df.columns:
                    df.rename(columns={'open_interest': '持仓量'}, inplace=True)
                df['日期'] = pd.to_datetime(df['日期'])
                df = df.sort_values('日期').dropna(subset=['收盘'])
                if len(df) < 10:
                    continue
                s_close = df.set_index('日期')['收盘'].astype(float)
                first = float(s_close.dropna().iloc[0])
                if first <= 0:
                    continue
                closes[u] = s_close / first * 100.0
                vols[u] = df.set_index('日期')['成交量'].astype(float)
                if '持仓量' in df.columns:
                    ois[u] = df.set_index('日期')['持仓量'].astype(float)
            except Exception:
                continue

        if not closes:
            print("商品主力等权合成: 无有效品种数据")
            return None

        panel_c = pd.DataFrame(closes).sort_index().ffill()
        for col in list(panel_c.columns):
            if panel_c[col].dropna().empty:
                panel_c = panel_c.drop(columns=[col])
        if panel_c.empty:
            return None
        eq_close = panel_c.mean(axis=1, skipna=True)
        sum_vol = pd.DataFrame(vols).sort_index().fillna(0).sum(axis=1)
        sum_oi = None
        if ois:
            sum_oi = pd.DataFrame(ois).sort_index().fillna(0).sum(axis=1)
        out = pd.DataFrame({
            '日期': eq_close.index,
            '收盘': eq_close.values,
            '开盘': eq_close.values,
            '最高': eq_close.values,
            '最低': eq_close.values,
            '成交量': sum_vol.reindex(eq_close.index).fillna(0).values,
            '成交额': 0.0,
            '持仓量': sum_oi.reindex(eq_close.index).fillna(0).values if sum_oi is not None else 0.0,
        })
        return out.dropna(subset=['收盘'])


# ============================================
# 指标计算模块
# ============================================

class IndicatorCalculator:
    """指标计算类"""
    
    def __init__(self, observation_end_date=None, weekly=False):
        self.observation_end_date = observation_end_date
        self.weekly = weekly
        self.fetcher = DataFetcher(observation_end_date)

    def _pct_window(self):
        return (52, 10) if self.weekly else (252, 20)

    def percentile_series(self, series):
        w, m = self._pct_window()
        return self.calculate_percentile_with_days(series, window=w, min_days=m)
    
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
        """获取市场情绪指标（日度或周度）"""
        indicators = {}
        
        index_returns_map = {
            '沪深300': '000300.XSHG',
            '中证A500': '000510.XSHG',
            '中证500': '000905.XSHG',
            '中证1000': '000852.XSHG',
            '国证2000': '399303.XSHE',
        }
        
        if self.weekly:
            for name, code in index_returns_map.items():
                try:
                    wdf = self.fetcher.get_index_weekly(code, approx_weeks=6)
                    if wdf is not None and not wdf.empty and len(wdf) >= 2:
                        wdf = wdf.sort_values('日期')
                        w_ret = (wdf['收盘'].iloc[-1] / wdf['收盘'].iloc[-2] - 1) * 100
                        indicators[f'{name}_daily_return'] = round(w_ret, 2)
                except Exception as e:
                    print(f"获取{name}周涨跌幅失败: {e}")
            try:
                widx = self.fetcher.get_index_weekly("000001.XSHG", approx_weeks=30)
                if widx is not None and not widx.empty:
                    widx = widx.sort_values('日期')
                    last_amt = widx['成交额'].iloc[-1] / 1e8
                    indicators['market_amount_20d_avg'] = round(last_amt, 2)
                    if len(widx) >= 5:
                        ch = (widx['成交额'].iloc[-1] / widx['成交额'].iloc[-5] - 1) * 100
                        indicators['market_amount_change_20d'] = round(ch, 2)
            except Exception as e:
                print(f"计算周度市场成交额失败: {e}")
        else:
            for name, code in index_returns_map.items():
                try:
                    df = self.fetcher.get_index_daily(code, period=2)
                    if df is not None and not df.empty and len(df) >= 2:
                        df = df.sort_values('日期')
                        daily_return = (df['收盘'].iloc[-1] / df['收盘'].iloc[-2] - 1) * 100
                        indicators[f'{name}_daily_return'] = round(daily_return, 2)
                except Exception as e:
                    print(f"获取{name}当日涨跌幅失败: {e}")
            try:
                df_index = self.fetcher.get_index_daily("000001.XSHG", period=40)
                if df_index is not None and not df_index.empty:
                    df_index = df_index.sort_values('日期')
                    avg_amount_20 = df_index['成交额'].tail(20).mean() / 1e8
                    indicators['market_amount_20d_avg'] = round(avg_amount_20, 2)
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
                end_date = self.fetcher._end(None)
                
                # 获取指数成分股
                try:
                    stocks = get_index_stocks(code, date=end_date)
                except TypeError:
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
                        
            except Exception as e:
                print(f"获取{name}估值数据失败: {e}")
        
        # 成长/价值比价（日度：近20交易日；周度：近5周）
        try:
            if self.weekly:
                df_cy = self.fetcher.get_index_weekly("399006.XSHE", approx_weeks=8)
                df_hs300 = self.fetcher.get_index_weekly("000300.XSHG", approx_weeks=8)
            else:
                df_cy = self.fetcher.get_index_daily("399006.XSHE", period=20)
                df_hs300 = self.fetcher.get_index_daily("000300.XSHG", period=20)
            if df_cy is not None and df_hs300 is not None:
                df_cy = df_cy.sort_values('日期')
                df_hs300 = df_hs300.sort_values('日期')
                if self.weekly:
                    n = min(5, len(df_cy), len(df_hs300))
                    df_cy = df_cy.tail(n)
                    df_hs300 = df_hs300.tail(n)
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
            end_date = self.fetcher._end(None)
            
            # 获取中证全指成分股作为全A代表
            try:
                stocks = get_index_stocks('000985.XSHG', date=end_date)
            except TypeError:
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
            
            history_period = 380 if self.weekly else 253
            df_prices = get_price(valid_stocks, count=history_period, end_date=end_date, frequency='daily', fields=['close'], panel=False)
            
            if df_prices is None or df_prices.empty:
                print("获取个股价格数据失败")
                return indicators
            
            # 透视为 日期×股票 的收盘价矩阵
            price_matrix = df_prices.pivot(index='time', columns='code', values='close')
            price_matrix = price_matrix.sort_index()
            
            if self.weekly:
                price_matrix = price_matrix.resample('W-FRI').last().dropna(how='all')
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
            
            latest_divergence = divergence_series.iloc[-1]
            indicators['market_divergence_daily'] = round(latest_divergence, 4)
            
            cum_n = 12 if self.weekly else 20
            if len(divergence_series) >= cum_n:
                divergence_20d = divergence_series.tail(cum_n).sum()
                indicators['market_divergence_20d'] = round(divergence_20d, 4)
            
            min_hist = cum_n * 2 if self.weekly else 40
            if len(divergence_series) >= min_hist:
                roll_win = 12 if self.weekly else 20
                rolling_div = divergence_series.rolling(window=roll_win).sum().dropna()
                percentile, days = self.percentile_series(rolling_div)
                if percentile is not None:
                    indicators['market_divergence_percentile'] = percentile
                    indicators['market_divergence_percentile_days'] = days
            
            # 不再内置「一九指数」主观阈值打分；仅输出分化度序列，评分侧用其历史分位
            
        except Exception as e:
            print(f"计算一九行情指标失败: {e}")
        
        return indicators
    
    def get_quant_indicators(self):
        """获取量化多头指标"""
        indicators = {}
        
        # 1. 交投活跃度
        try:
            if self.weekly:
                df_index = self.fetcher.get_index_weekly("000001.XSHG", approx_weeks=30)
                if df_index is not None and not df_index.empty:
                    df_index = df_index.sort_values('日期')
                    avg_amount = df_index['成交额'].tail(4).mean() / 1e8
                    indicators['market_amount_20d_avg'] = round(avg_amount, 2)
            else:
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
            if self.weekly:
                df_all = self.fetcher.get_index_weekly("000985.XSHG", approx_weeks=8)
                if df_all is not None and not df_all.empty:
                    total_amount = df_all['成交额'].sum()
                    for name, code in index_symbols.items():
                        df_idx = self.fetcher.get_index_weekly(code, approx_weeks=8)
                        if df_idx is not None and not df_idx.empty:
                            idx_amount = df_idx['成交额'].sum()
                            indicators[f'{name}_crowding'] = round(idx_amount / total_amount * 100, 2)
            else:
                df_all = self.fetcher.get_index_daily("000985.XSHG", period=20)
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
        """CTA：商品主力等权合成指数 — 年化波动、周均成交、价格水平及其历史分位。"""
        indicators = {}
        try:
            daily = self.fetcher.get_commodity_equal_weight_main_daily(period=800)
            if daily is None or daily.empty:
                print("CTA: 商品主力等权合成数据不足")
                return indicators
            daily = daily.sort_values('日期')
            if self.weekly:
                df = _resample_ohlcv_to_weekly_friday(daily)
            else:
                df = daily
            if df is None or df.empty or len(df) < 30:
                print("CTA: 重采样后数据不足")
                return indicators
            df = df.sort_values('日期')
            ret = df['收盘'].pct_change()
            win = 4 if self.weekly else 20
            ann = 52 if self.weekly else 252
            vol_s = ret.rolling(win).std() * np.sqrt(ann) * 100
            cur_vol = vol_s.iloc[-1]
            if not np.isnan(cur_vol):
                indicators['commodity_eq_volatility'] = round(float(cur_vol), 2)
            vp, vd = self.percentile_series(vol_s.dropna())
            if vp is not None:
                indicators['commodity_eq_volatility_percentile'] = vp
                indicators['commodity_eq_volatility_percentile_days'] = vd
            avg_vol = df['成交量'].tail(win).mean()
            if not np.isnan(avg_vol):
                indicators['commodity_eq_avg_volume'] = round(float(avg_vol), 2)
            ap, ad = self.percentile_series(df['成交量'])
            if ap is not None:
                indicators['commodity_eq_volume_percentile'] = ap
                indicators['commodity_eq_volume_percentile_days'] = ad
            lvl = df['收盘'].iloc[-1]
            if not np.isnan(lvl):
                indicators['commodity_eq_price_level'] = round(float(lvl), 4)
            pp, pd_ = self.percentile_series(df['收盘'])
            if pp is not None:
                indicators['commodity_eq_price_percentile'] = pp
                indicators['commodity_eq_price_percentile_days'] = pd_
            if '持仓量' in df.columns:
                oi = df['持仓量'].iloc[-1]
                if not np.isnan(oi):
                    indicators['commodity_eq_open_interest'] = round(float(oi), 2)
                op, od = self.percentile_series(df['持仓量'])
                if op is not None:
                    indicators['commodity_eq_oi_percentile'] = op
                    indicators['commodity_eq_oi_percentile_days'] = od
        except Exception as e:
            print(f"计算CTA商品等权指标失败: {e}")
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
                if self.weekly:
                    df_etf = self.fetcher.get_etf_weekly(etf_code, approx_weeks=140)
                else:
                    df_etf = self.fetcher.get_etf_daily(etf_code, period=252)
                if df_etf is None or df_etf.empty:
                    continue
                win = 4 if self.weekly else 20
                ann = 52 if self.weekly else 252
                avg_amt = df_etf['成交额'].tail(win).mean() / 1e8
                indicators[f'{etf_name}_avg_amount_20d'] = round(avg_amt, 2)
                ap, ad = self.percentile_series(df_etf['成交额'])
                if ap is not None:
                    indicators[f'{etf_name}_amount_percentile'] = ap
                    indicators[f'{etf_name}_amount_percentile_days'] = ad
                df_etf = df_etf.sort_values('日期')
                df_etf['daily_return'] = df_etf['收盘'].pct_change()
                df_etf['volatility_20d'] = df_etf['daily_return'].rolling(win).std() * np.sqrt(ann) * 100
                current_vol = df_etf['volatility_20d'].iloc[-1]
                if not np.isnan(current_vol):
                    indicators[f'{etf_name}_volatility_20d'] = round(float(current_vol), 2)
                vp, vd = self.percentile_series(df_etf['volatility_20d'])
                if vp is not None:
                    indicators[f'{etf_name}_volatility_percentile'] = vp
                    indicators[f'{etf_name}_volatility_percentile_days'] = vd
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
                if self.weekly:
                    df_future = self.fetcher.get_futures_main_weekly(future_code, approx_weeks=140)
                    win, ann, min_bars = 4, 52, 25
                else:
                    df_future = self.fetcher.get_futures_main_contract(future_code, period=252)
                    win, ann, min_bars = 20, 252, 20
                if df_future is not None and not df_future.empty and len(df_future) >= min_bars:
                    df_future = df_future.sort_values('日期')
                    df_future['daily_return'] = df_future['收盘价'].pct_change()
                    df_future['volatility_20d'] = df_future['daily_return'].rolling(win).std() * np.sqrt(ann) * 100
                    current_vol = df_future['volatility_20d'].iloc[-1]
                    if not np.isnan(current_vol):
                        indicators[f'{future_code}_volatility_20d'] = round(float(current_vol), 2)
                    vol_percentile, vol_days = self.percentile_series(df_future['volatility_20d'])
                    if vol_percentile is not None:
                        indicators[f'{future_code}_volatility_percentile'] = vol_percentile
                        indicators[f'{future_code}_volatility_percentile_days'] = vol_days
                    amount_percentile, amount_days = self.percentile_series(df_future['成交额'])
                    if amount_percentile is not None:
                        indicators[f'{future_code}_amount_percentile'] = amount_percentile
                        indicators[f'{future_code}_amount_percentile_days'] = amount_days
            except Exception:
                continue
        
        # ===== 3. 期权套利指标 =====
        # 动态获取期权合约的隐含波动率数据
        
        def get_option_iv(underlying_code, option_type='ETF', name=''):
            """获取期权的隐含波动率 - 扩大扫描范围"""
            try:
                end_date = self.fetcher._end(None)
                wk = self.weekly
                iv_count = 400 if wk else 252
                iv_window, iv_min = ((52, 10) if wk else (252, 20))
                
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
                        df_option = get_price(best_code, count=iv_count, end_date=end_date, frequency='daily', fields=['close', 'implied_volatility'])
                        if df_option is not None and not df_option.empty:
                            df_option = df_option.reset_index()
                            df_option.rename(columns={'index': '日期', 'implied_volatility': 'IV'}, inplace=True)
                            df_option['日期'] = pd.to_datetime(df_option['日期'])
                            df_option = df_option.sort_values('日期')
                            df_option = df_option.dropna(subset=['IV'])
                            if wk:
                                df_option = df_option.set_index('日期')['IV'].resample('W-FRI').last().dropna().reset_index()
                                df_option.columns = ['日期', 'IV']
                            if len(df_option) >= iv_min:
                                current_iv = df_option['IV'].iloc[-1]
                                indicators[f'{name}_IV'] = round(current_iv * 100, 2)
                                
                                iv_percentile, iv_days = self.calculate_percentile_with_days(
                                    df_option['IV'], window=iv_window, min_days=iv_min)
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
                        
                        df_option = get_price(best_code, count=iv_count, end_date=end_date, frequency='daily', fields=['close', 'implied_volatility'])
                        if df_option is not None and not df_option.empty:
                            df_option = df_option.reset_index()
                            df_option.rename(columns={'index': '日期', 'implied_volatility': 'IV'}, inplace=True)
                            df_option['日期'] = pd.to_datetime(df_option['日期'])
                            df_option = df_option.sort_values('日期')
                            df_option = df_option.dropna(subset=['IV'])
                            if wk:
                                df_option = df_option.set_index('日期')['IV'].resample('W-FRI').last().dropna().reset_index()
                                df_option.columns = ['日期', 'IV']
                            if len(df_option) >= iv_min:
                                current_iv = df_option['IV'].iloc[-1]
                                indicators[f'{name}_IV'] = round(current_iv * 100, 2)
                                
                                iv_percentile, iv_days = self.calculate_percentile_with_days(
                                    df_option['IV'], window=iv_window, min_days=iv_min)
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
                if self.weekly:
                    df_future = self.fetcher.get_futures_main_weekly(future_code, approx_weeks=8)
                    df_index = self.fetcher.get_index_weekly(index_code, approx_weeks=8)
                else:
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
        """获取市场中性策略指标（不含市场形态/哑铃类指标）。"""
        indicators = {}
        
        # 1. 交投活跃度
        try:
            if self.weekly:
                df_index = self.fetcher.get_index_weekly("000001.XSHG", approx_weeks=30)
                if df_index is not None and not df_index.empty:
                    df_index = df_index.sort_values('日期')
                    avg_amount = df_index['成交额'].tail(4).mean() / 1e8
                    indicators['market_amount_20d_avg'] = round(avg_amount, 2)
            else:
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
            if self.weekly:
                df_all = self.fetcher.get_index_weekly("000985.XSHG", approx_weeks=8)
                if df_all is not None and not df_all.empty:
                    total_amount = df_all['成交额'].sum()
                    for name, code in index_symbols.items():
                        df_idx = self.fetcher.get_index_weekly(code, approx_weeks=8)
                        if df_idx is not None and not df_idx.empty:
                            idx_amount = df_idx['成交额'].sum()
                            indicators[f'{name}_crowding'] = round(idx_amount / total_amount * 100, 2)
            else:
                df_all = self.fetcher.get_index_daily("000985.XSHG", period=20)
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
                if self.weekly:
                    df_future = self.fetcher.get_futures_main_weekly(future_code, approx_weeks=8)
                    df_index = self.fetcher.get_index_weekly(index_code, approx_weeks=8)
                else:
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
        
        return indicators


# ============================================
# 策略评分模块
# ============================================

class StrategyScorer:
    """策略评分：以各指标在自身历史分布中的分位数为主，线性映射到加减分（不依赖 SCORE_WEIGHTS 占位）。"""
    
    def __init__(self, observation_end_date=None, weekly=False):
        self.calculator = IndicatorCalculator(observation_end_date, weekly=weekly)

    @staticmethod
    def _clamp_score(s):
        return max(0, min(100, int(round(s))))

    def _pct_contrib(self, pct, days, label, weight, details):
        """分位数 P∈[0,100] 相对中位 50 线性贡献，权重为极端满分幅度。"""
        if pct is None:
            details.append(f"{label}: 历史分位数据不足")
            return 0.0
        c = (float(pct) - 50.0) / 50.0 * float(weight)
        unit = "周" if self.calculator.weekly else "天"
        dstr = f"{days}{unit}" if days else ""
        details.append(f"{label} 历史分位{pct:.1f}% ({dstr}): {c:+.1f}")
        return c

    def _series_market_amount_20d_avg(self, count=300):
        w = self.calculator.weekly
        if w:
            df = self.calculator.fetcher.get_index_weekly("000001.XSHG", approx_weeks=count // 5)
            if df is None or df.empty or len(df) < 10:
                return None
            df = df.sort_values('日期')
            return (df['成交额'].rolling(4).mean() / 1e8).dropna()
        df = self.calculator.fetcher.get_index_daily("000001.XSHG", period=count)
        if df is None or df.empty or len(df) < 40:
            return None
        df = df.sort_values('日期')
        return (df['成交额'].rolling(20).mean() / 1e8).dropna()

    def _series_amount_change_20d(self, count=300):
        w = self.calculator.weekly
        if w:
            df = self.calculator.fetcher.get_index_weekly("000001.XSHG", approx_weeks=count // 5)
            if df is None or df.empty or len(df) < 10:
                return None
            df = df.sort_values('日期')
            a = df['成交额']
            return ((a / a.shift(4) - 1.0) * 100.0).dropna()
        df = self.calculator.fetcher.get_index_daily("000001.XSHG", period=count)
        if df is None or df.empty or len(df) < 40:
            return None
        df = df.sort_values('日期')
        a = df['成交额']
        return ((a / a.shift(20) - 1.0) * 100.0).dropna()

    def _series_growth_value_ratio(self, count=120):
        w = self.calculator.weekly
        if w:
            df_cy = self.calculator.fetcher.get_index_weekly("399006.XSHE", approx_weeks=count)
            df_hs = self.calculator.fetcher.get_index_weekly("000300.XSHG", approx_weeks=count)
        else:
            df_cy = self.calculator.fetcher.get_index_daily("399006.XSHE", period=count)
            df_hs = self.calculator.fetcher.get_index_daily("000300.XSHG", period=count)
        if df_cy is None or df_hs is None or df_cy.empty or df_hs.empty:
            return None
        df_cy = df_cy.sort_values('日期').set_index('日期')['收盘']
        df_hs = df_hs.sort_values('日期').set_index('日期')['收盘']
        m = pd.concat([df_cy, df_hs], axis=1, join='inner').dropna()
        lag = 4 if w else 20
        if len(m) < lag + 5:
            return None
        r_cy = m.iloc[:, 0].pct_change(lag) * 100.0
        r_hs = m.iloc[:, 1].pct_change(lag) * 100.0
        return (r_cy - r_hs).dropna()

    def _series_crowding(self, index_code, count=300):
        w = self.calculator.weekly
        if w:
            df_all = self.calculator.fetcher.get_index_weekly("000985.XSHG", approx_weeks=count // 5)
            df_idx = self.calculator.fetcher.get_index_weekly(index_code, approx_weeks=count // 5)
        else:
            df_all = self.calculator.fetcher.get_index_daily("000985.XSHG", period=count)
            df_idx = self.calculator.fetcher.get_index_daily(index_code, period=count)
        if df_all is None or df_idx is None or df_all.empty or df_idx.empty:
            return None
        df_all = df_all.sort_values('日期').set_index('日期')['成交额']
        df_idx = df_idx.sort_values('日期').set_index('日期')['成交额']
        m = pd.concat([df_all, df_idx], axis=1, join='inner').dropna()
        win = 4 if w else 20
        if len(m) < win + 5:
            return None
        roll_a = m.iloc[:, 0].rolling(win).sum()
        roll_i = m.iloc[:, 1].rolling(win).sum()
        return (roll_i / roll_a * 100.0).dropna()

    def _contrib_from_series(self, series, label, weight, details):
        _, min_pts = self.calculator._pct_window()
        if series is None or len(series) < min_pts:
            details.append(f"{label}: 序列数据不足")
            return 0.0
        pct, days = self.calculator.percentile_series(series)
        return self._pct_contrib(pct, days, label, weight, details)

    def score_subjective_long(self):
        score, details = 50.0, []
        try:
            w = self.calculator.weekly
            score += self._contrib_from_series(
                self._series_market_amount_20d_avg(),
                '上证周均成交额(亿元,4周)' if w else '上证20日均成交额(亿元)',
                16, details)
            score += self._contrib_from_series(
                self._series_amount_change_20d(),
                '上证成交额4周环比(%)' if w else '上证成交额20日环比(%)',
                12, details)
            score += self._contrib_from_series(
                self._series_growth_value_ratio(),
                '创业板相对沪深300(4周超额%)' if w else '创业板相对沪深300(20日超额%)',
                12, details)
        except Exception as e:
            details.append(f"评分计算出错: {e}")
        return {'score': self._clamp_score(score), 'details': details, 'level': self._get_level(score), 'trend': '→ 持平'}

    def score_quant_long(self):
        score, details = 50.0, []
        try:
            q = self.calculator.get_quant_indicators()
            w = self.calculator.weekly
            score += self._contrib_from_series(
                self._series_market_amount_20d_avg(),
                '市场周均成交额(亿元,4周)' if w else '市场20日均成交额(亿元)',
                12, details)
            score += self._contrib_from_series(
                self._series_crowding('000852.XSHG'),
                '中证1000拥挤度(4周)' if w else '中证1000拥挤度(占全市场%)',
                18, details)
            div_pct = q.get('market_divergence_percentile')
            div_days = q.get('market_divergence_percentile_days', 0)
            score += self._pct_contrib(
                div_pct, div_days,
                '分化度(12周累计)历史分位' if w else '市值加权-等权分化度(20日累计)历史分位',
                22, details)
        except Exception as e:
            details.append(f"评分计算出错: {e}")
        return {'score': self._clamp_score(score), 'details': details, 'level': self._get_level(score), 'trend': '→ 持平'}

    def score_cta(self):
        score, details = 50.0, []
        try:
            c = self.calculator.get_cta_indicators()
            w = self.calculator.weekly
            lab_vol = '商品等权指数年化波动分位' + ('(4周)' if w else '(20日)')
            lab_volm = '商品等权指数周均成交量分位' if w else '商品等权指数成交量分位'
            lab_px = '商品等权指数价位分位'
            vp, vd = c.get('commodity_eq_volatility_percentile'), c.get('commodity_eq_volatility_percentile_days')
            score += self._pct_contrib(vp, vd, lab_vol, 22, details)
            ap, ad = c.get('commodity_eq_volume_percentile'), c.get('commodity_eq_volume_percentile_days')
            score += self._pct_contrib(ap, ad, lab_volm, 18, details)
            pp, pd_ = c.get('commodity_eq_price_percentile'), c.get('commodity_eq_price_percentile_days')
            score += self._pct_contrib(pp, pd_, lab_px, 18, details)
            op, od = c.get('commodity_eq_oi_percentile'), c.get('commodity_eq_oi_percentile_days')
            score += self._pct_contrib(op, od, '商品等权指数持仓量分位', 14, details)
        except Exception as e:
            details.append(f"评分计算出错: {e}")
        return {'score': self._clamp_score(score), 'details': details, 'level': self._get_level(score), 'trend': '→ 持平'}

    def score_etf_arbitrage(self):
        score, details = 50.0, []
        try:
            a = self.calculator.get_arbitrage_indicators()
            etfs = ['300ETF', '500ETF', '1000ETF', '50ETF', '创业板ETF', '科创50ETF']
            vp, vd, ap, ad = [], [], [], []
            w = self.calculator.weekly
            for e in etfs:
                p = a.get(f'{e}_volatility_percentile')
                if p is not None:
                    vp.append(p)
                    vd.append(a.get(f'{e}_volatility_percentile_days', 0))
                p = a.get(f'{e}_amount_percentile')
                if p is not None:
                    ap.append(p)
                    ad.append(a.get(f'{e}_amount_percentile_days', 0))
            if vp:
                score += self._pct_contrib(
                    float(np.mean(vp)), int(np.mean(vd)) if vd else 0,
                    'ETF波动率分位(均值' + ('，周线)' if w else ')'),
                    22, details)
            else:
                details.append('ETF波动率分位: 数据不足')
            if ap:
                score += self._pct_contrib(
                    float(np.mean(ap)), int(np.mean(ad)) if ad else 0,
                    'ETF成交额分位(均值' + ('，周线)' if w else ')'),
                    20, details)
            else:
                details.append('ETF成交额分位: 数据不足')
        except Exception as e:
            details.append(f"评分计算出错: {e}")
        return {'score': self._clamp_score(score), 'details': details, 'level': self._get_level(score), 'trend': '→ 持平'}

    def score_index_arbitrage(self):
        score, details = 50.0, []
        try:
            a = self.calculator.get_arbitrage_indicators()
            vp, vd, ap, ad = [], [], [], []
            for f in ['IF', 'IC', 'IM']:
                p = a.get(f'{f}_volatility_percentile')
                if p is not None:
                    vp.append(p)
                    vd.append(a.get(f'{f}_volatility_percentile_days', 0))
                p = a.get(f'{f}_amount_percentile')
                if p is not None:
                    ap.append(p)
                    ad.append(a.get(f'{f}_amount_percentile_days', 0))
            if vp:
                score += self._pct_contrib(float(np.mean(vp)), int(np.mean(vd)) if vd else 0, '股指期货波动率分位(均值)', 22, details)
            else:
                details.append('股指期货波动率分位: 数据不足')
            if ap:
                score += self._pct_contrib(float(np.mean(ap)), int(np.mean(ad)) if ad else 0, '股指期货成交额分位(均值)', 18, details)
            else:
                details.append('股指期货成交额分位: 数据不足')
        except Exception as e:
            details.append(f"评分计算出错: {e}")
        return {'score': self._clamp_score(score), 'details': details, 'level': self._get_level(score), 'trend': '→ 持平'}

    def score_option_arbitrage(self):
        score, details = 50.0, []
        try:
            a = self.calculator.get_arbitrage_indicators()
            names = (['50ETF', '300ETF', '500ETF', '科创50ETF', '创业板ETF', '深300ETF']
                     + ['300股指', '1000股指', '50股指'])
            pcts, days = [], []
            for n in names:
                p = a.get(f'{n}_IV_percentile')
                if p is not None:
                    pcts.append(p)
                    days.append(a.get(f'{n}_IV_percentile_days', 0))
            if pcts:
                w = self.calculator.weekly
                score += self._pct_contrib(
                    float(np.mean(pcts)), int(np.mean(days)) if days else 0,
                    '期权IV分位(均值' + ('，周线)' if w else ')'),
                    26, details)
                if len(pcts) >= 2:
                    diff = max(pcts) - min(pcts)
                    bonus = min(14.0, diff / 50.0 * 14.0)
                    score += bonus
                    details.append(f"IV分位跨标的分化(极差{diff:.1f}%): +{bonus:.1f}")
            else:
                details.append('期权IV分位: 数据不足')
        except Exception as e:
            details.append(f"评分计算出错: {e}")
        return {'score': self._clamp_score(score), 'details': details, 'level': self._get_level(score), 'trend': '→ 持平'}

    def score_market_neutral(self):
        score, details = 50.0, []
        try:
            w = self.calculator.weekly
            n = self.calculator.get_market_neutral_indicators()
            score += self._contrib_from_series(
                self._series_market_amount_20d_avg(),
                '市场周均成交额(亿元,4周)' if w else '市场20日均成交额(亿元)',
                12, details)
            score += self._contrib_from_series(
                self._series_crowding('000905.XSHG'),
                '中证500拥挤度(4周)' if w else '中证500拥挤度(占全市场%)',
                16, details)
            ic = n.get('IC_basis_annual')
            im = n.get('IM_basis_annual')
            ser_mix = self._series_ic_im_annual_basis_avg(count=520 if w else 320)
            _, min_pts = self.calculator._pct_window()
            if ser_mix is not None and len(ser_mix) >= min_pts:
                pct, days = self.calculator.percentile_series(ser_mix)
                score += self._pct_contrib(
                    pct, days,
                    'IC+IM年化基差(周线)历史分位' if w else 'IC+IM年化基差(逐日合成)历史分位',
                    24, details)
            elif ic is not None and im is not None:
                details.append(f"IC+IM年化基差当前({(ic+im)/2:.2f}%): 历史序列不足，未打分位")
            else:
                details.append('IC/IM年化基差: 数据不足')
        except Exception as e:
            details.append(f"评分计算出错: {e}")
        return {'score': self._clamp_score(score), 'details': details, 'level': self._get_level(score), 'trend': '→ 持平'}

    def _series_futures_annual_basis(self, future_code, index_code, count=320):
        """(期货收-指数收)/指数收，年化基差近似序列（周度为周五对齐后逐周）。"""
        try:
            w = self.calculator.weekly
            if w:
                df_f = self.calculator.fetcher.get_futures_main_weekly(
                    future_code, approx_weeks=max(140, count // 5))
                df_i = self.calculator.fetcher.get_index_weekly(
                    index_code, approx_weeks=max(140, count // 5))
            else:
                df_f = self.calculator.fetcher.get_futures_main_contract(future_code, period=count)
                df_i = self.calculator.fetcher.get_index_daily(index_code, period=count)
            if df_f is None or df_i is None or df_f.empty or df_i.empty:
                return None
            df_f = df_f.sort_values('日期').set_index('日期')['收盘价']
            df_i = df_i.sort_values('日期').set_index('日期')['收盘']
            m = pd.concat([df_f, df_i], axis=1, join='inner').dropna()
            min_pts = self.calculator._pct_window()[1]
            if len(m) < min_pts:
                return None
            raw = (m.iloc[:, 0] / m.iloc[:, 1] - 1.0) * 100.0 * 4.0
            return raw.dropna()
        except Exception:
            return None

    def _series_ic_im_annual_basis_avg(self, count=320):
        s_ic = self._series_futures_annual_basis('IC', '000905.XSHG', count)
        s_im = self._series_futures_annual_basis('IM', '000852.XSHG', count)
        if s_ic is None or s_im is None:
            return None
        m = pd.concat([s_ic, s_im], axis=1, join='inner').dropna()
        min_pts = self.calculator._pct_window()[1]
        if m.empty or len(m) < min_pts:
            return None
        return ((m.iloc[:, 0] + m.iloc[:, 1]) / 2.0)
    
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
    
    def __init__(self, output_path=None, observation_end_date=None, run_time_str=None, weekly=True):
        self.observation_end_date = observation_end_date
        self.run_time_str = run_time_str or dt.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        self.weekly = weekly
        if output_path is None:
            timestamp = dt.datetime.now().strftime('%Y%m%d_%H%M%S')
            # 聚宽环境中使用当前目录
            tag = 'weekly' if weekly else 'daily'
            self.output_path = 'strategy_monitor_{}_{}.xlsx'.format(tag, timestamp)
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
        
        title = '策略环境监测报告（周度）' if self.weekly else '策略环境监测报告（日度）'
        ws['A1'] = title
        ws['A1'].font = Font(bold=True, size=16, color='1F4E78')
        ws.merge_cells('A1:E1')
        ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
        
        obs = self.observation_end_date or ''
        ws['A2'] = '观测数据截止: {}  |  报告生成: {}'.format(obs, self.run_time_str)
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
        wk = self.weekly
        
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
        
        # ===== 第一部分：主要指数涨跌幅 =====
        sec_title = '【主要指数周涨跌幅】' if wk else '【主要指数当日涨跌幅】'
        ws.cell(row=row, column=1, value=sec_title)
        ws.cell(row=row, column=1).font = Font(bold=True, size=11, color='1F4E78')
        ws.merge_cells(f'A{row}:D{row}')
        row += 1
        
        index_list = ['沪深300', '中证A500', '中证500', '中证1000', '国证2000']
        for idx_name in index_list:
            ret = indicators.get(f'{idx_name}_daily_return')
            if ret is not None:
                lbl = '{}周涨跌幅'.format(idx_name) if wk else '{}当日涨跌幅'.format(idx_name)
                ws.cell(row=row, column=1, value=lbl)
                ws.cell(row=row, column=2, value=f"{ret}%")
                ws.cell(row=row, column=3, value='%')
                desc = '{}指数周涨跌幅（最近一周相对上一周）'.format(idx_name) if wk else '{}指数当日涨跌幅'.format(idx_name)
                ws.cell(row=row, column=4, value=desc)
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
            amt_lbl = '近4周上证周均成交额' if wk else '近20日市场平均成交额'
            amt_desc = '最近一周全市场成交额（亿元）' if wk else '反映市场整体活跃度'
            ws.cell(row=row, column=1, value=amt_lbl)
            ws.cell(row=row, column=2, value=f"{market_amount:.2f}")
            ws.cell(row=row, column=3, value='亿元')
            ws.cell(row=row, column=4, value=amt_desc)
            for col in range(1, 5):
                cell = ws.cell(row=row, column=col)
                cell.border = Border(left=Side(style='thin'), right=Side(style='thin'),
                                    top=Side(style='thin'), bottom=Side(style='thin'))
                cell.alignment = Alignment(horizontal='left' if col in [1, 4] else 'center', vertical='center')
            row += 1
        
        amount_change = indicators.get('market_amount_change_20d')
        if amount_change is not None:
            ch_lbl = '上证成交额4周环比' if wk else '近20日成交额变化率'
            ch_desc = '最近一周成交额相对4周前一周的变化率' if wk else '反映市场热度变化'
            ws.cell(row=row, column=1, value=ch_lbl)
            ws.cell(row=row, column=2, value=f"{amount_change}%")
            ws.cell(row=row, column=3, value='%')
            ws.cell(row=row, column=4, value=ch_desc)
            for col in range(1, 5):
                cell = ws.cell(row=row, column=col)
                cell.border = Border(left=Side(style='thin'), right=Side(style='thin'),
                                    top=Side(style='thin'), bottom=Side(style='thin'))
                cell.alignment = Alignment(horizontal='left' if col in [1, 4] else 'center', vertical='center')
            row += 1
        
        row += 1  # 空行
        
        # ===== 第三部分：估值指标（当前 PE/PB，不含历史分位数） =====
        ws.cell(row=row, column=1, value='【估值指标 - 当前 PE/PB】')
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
            pb = indicators.get(f'{idx_code}_PB')
            
            if pe is not None:
                ws.cell(row=row, column=1, value=f'{idx_name} PE')
                ws.cell(row=row, column=2, value=f"{pe:.2f}倍")
                ws.cell(row=row, column=3, value='倍')
                ws.cell(row=row, column=4, value=f'{idx_name}市盈率（成分加权，当前截面）')
                for col in range(1, 5):
                    cell = ws.cell(row=row, column=col)
                    cell.border = Border(left=Side(style='thin'), right=Side(style='thin'),
                                        top=Side(style='thin'), bottom=Side(style='thin'))
                    cell.alignment = Alignment(horizontal='left' if col in [1, 4] else 'center', vertical='center')
                row += 1
            
            if pb is not None:
                ws.cell(row=row, column=1, value=f'{idx_name} PB')
                ws.cell(row=row, column=2, value=f"{pb:.2f}倍")
                ws.cell(row=row, column=3, value='倍')
                ws.cell(row=row, column=4, value=f'{idx_name}市净率（成分加权，当前截面）')
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
            gv_desc = (
                '创业板指与沪深300指数近5周收益率差值，反映风格偏好' if wk
                else '创业板指与沪深300指数20日收益率差值，反映风格偏好')
            ws.cell(row=row, column=4, value=gv_desc)
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
        wk = self.weekly
        
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
        
        roll = '近约4周' if wk else '近20日'
        div_u = '周' if wk else '天'
        indicator_descriptions = {
            'market_amount_20d_avg': '{}市场平均成交额，反映市场活跃度'.format(roll),
            '沪深300_crowding': '沪深300指数成交额占全市场比例（{}），反映大盘拥挤度'.format(roll),
            '中证500_crowding': '中证500指数成交额占全市场比例（{}），反映中盘拥挤度'.format(roll),
            '中证1000_crowding': '中证1000指数成交额占全市场比例（{}），反映小盘拥挤度'.format(roll),
            'market_divergence_daily': '市值加权收益与等权收益的差值（{}度），正值表示大盘股跑赢多数个股'.format('周' if wk else '日'),
            'market_divergence_20d': '{}分化度累计值，持续正值表示一九行情持续'.format('12周' if wk else '近20日'),
            'market_divergence_percentile': '{}累计分化度在历史分布中的分位数'.format('12周' if wk else '20日'),
            'market_divergence_percentile_days': '计算分化度分位数所用的历史{}'.format(div_u),
        }
        
        indicator_units = {
            'market_amount_20d_avg': '亿元',
            '沪深300_crowding': '%',
            '中证500_crowding': '%',
            '中证1000_crowding': '%',
            'market_divergence_daily': '%',
            'market_divergence_20d': '%',
            'market_divergence_percentile': '%',
            'market_divergence_percentile_days': div_u,
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
        """CTA：商品主力等权合成（日度/周度与 IndicatorCalculator 一致）"""
        ws = wb.create_sheet("CTA策略指标")
        wk = self.weekly
        u = '周' if wk else '日'
        win_label = '4周滚动' if wk else '20日滚动'

        ws['A1'] = 'CTA策略 - 商品主力等权合成指标'
        ws['A1'].font = Font(bold=True, size=14, color='1F4E78')
        ws.merge_cells('A1:E1')
        ws['A1'].alignment = Alignment(horizontal='center', vertical='center')

        headers = ['标的', '指标名称', '当前值', '历史分位数', '单位']
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
        rows_spec = [
            ('合成指数', '{}年化波动率'.format(win_label), 'commodity_eq_volatility', '%',
             'commodity_eq_volatility_percentile', 'commodity_eq_volatility_percentile_days'),
            ('合成指数', '{}均成交量'.format('近4周' if wk else '近20日'), 'commodity_eq_avg_volume', '手',
             'commodity_eq_volume_percentile', 'commodity_eq_volume_percentile_days'),
            ('合成指数', '价格水平(归一化后等权收盘)', 'commodity_eq_price_level', '点',
             'commodity_eq_price_percentile', 'commodity_eq_price_percentile_days'),
            ('合成指数', '周末/日末总持仓量' if wk else '日末总持仓量', 'commodity_eq_open_interest', '手',
             'commodity_eq_oi_percentile', 'commodity_eq_oi_percentile_days'),
        ]

        for who, iname, vkey, unit, pkey, dkey in rows_spec:
            val = indicators.get(vkey)
            if val is None and indicators.get(pkey) is None:
                continue
            ws.cell(row=row, column=1, value=who)
            ws.cell(row=row, column=2, value=iname)
            ws.cell(row=row, column=3, value='-' if val is None else val)
            pct = indicators.get(pkey)
            nd = indicators.get(dkey, 0)
            if pct is not None:
                ws.cell(row=row, column=4, value='{}% (基于{}{})'.format(pct, nd, u))
            else:
                ws.cell(row=row, column=4, value='数据不足')
            ws.cell(row=row, column=5, value=unit)
            for col in range(1, 6):
                cell = ws.cell(row=row, column=col)
                cell.border = Border(
                    left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin')
                )
                cell.alignment = Alignment(horizontal='center', vertical='center')
            row += 1

        ws.column_dimensions['A'].width = 14
        ws.column_dimensions['B'].width = 28
        ws.column_dimensions['C'].width = 18
        ws.column_dimensions['D'].width = 24
        ws.column_dimensions['E'].width = 8

    def _create_arbitrage_sheet(self, wb, indicators):
        """创建套利策略指标sheet - 优化展示格式"""
        ws = wb.create_sheet("套利策略指标")
        wk = self.weekly
        u = '周' if wk else '天'
        
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
                amt_nm = '近4周平均周成交额' if wk else '近20日平均成交额'
                ws.cell(row=row, column=2, value=amt_nm)
                ws.cell(row=row, column=3, value=f"{amount:.2f}")
                if amount_pct is not None:
                    pct_display = f"{amount_pct}% (基于{amount_days}{u})" if amount_days > 0 else f"{amount_pct}%"
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
                vol_nm = '4周年化波动率' if wk else '20日年化波动率'
                ws.cell(row=row, column=2, value=vol_nm)
                ws.cell(row=row, column=3, value=f"{vol}%")
                if vol_pct is not None:
                    pct_display = f"{vol_pct}% (基于{vol_days}{u})" if vol_days > 0 else f"{vol_pct}%"
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
                vol_nm = '4周年化波动率' if wk else '20日年化波动率'
                ws.cell(row=row, column=2, value=vol_nm)
                ws.cell(row=row, column=3, value=f"{vol}%")
                if vol_pct is not None:
                    pct_display = f"{vol_pct}% (基于{vol_days}{u})" if vol_days > 0 else f"{vol_pct}%"
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
                pct_display = f"{amt_pct}% (基于{amt_days}{u})" if amt_days > 0 else f"{amt_pct}%"
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
                    pct_display = f"{iv_pct}% (基于{iv_days}{u})" if iv_days > 0 else f"{iv_pct}%"
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
        wk = self.weekly
        
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
            amt_lbl = '近4周市场周均成交额' if wk else '近20日市场平均成交额'
            ws.cell(row=row, column=1, value=amt_lbl)
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
                roll_desc = '近约4周成交额占全市场比例' if wk else '近20日成交额占全市场比例'
                ws.cell(row=row, column=4, value='{}；{}'.format(idx_name, roll_desc))
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

def collect_all_indicators(observation_end_date=None, weekly=True):
    """收集所有策略的指标；observation_end_date 为数据截止日（周度监测用上个周五对齐交易日）。"""
    print("=" * 60)
    print("正在收集各策略指标...")
    print("=" * 60)
    
    calculator = IndicatorCalculator(observation_end_date, weekly=weekly)
    all_indicators = {'_meta': {'observation_end_date': observation_end_date or '', 'weekly': weekly}}
    
    # 1. 主观多头指标
    print("\n[1/7] 收集主观多头指标...")
    subjective = {}
    subjective.update(calculator.get_sentiment_indicators())
    subjective.update(calculator.get_valuation_indicators())
    all_indicators['subjective'] = subjective
    print(f"  - 市场情绪指标: {len([k for k in subjective.keys() if 'amount' in k])}个")
    print(f"  - 估值指标(PE/PB): {len([k for k in subjective.keys() if k.endswith('_PE') or k.endswith('_PB')])}个")
    
    # 2. 量化多头指标
    print("\n[2/7] 收集量化多头指标...")
    quant = calculator.get_quant_indicators()
    all_indicators['quant'] = quant
    print(f"  - 拥挤度指标: {len([k for k in quant.keys() if 'crowding' in k])}个")
    print(f"  - 分化度相关指标: {len([k for k in quant.keys() if 'divergence' in k])}个")
    if quant.get('market_divergence_daily') is not None:
        lbl = '当周' if weekly else '当日'
        print("  - {}市值加权-等权分化度: {}".format(lbl, quant.get('market_divergence_daily')))
    
    # 3. CTA指标（仅商品期货）
    print("\n[3/7] 收集CTA策略指标（商品期货）...")
    cta = calculator.get_cta_indicators()
    all_indicators['cta'] = cta
    print("  - 商品等权合成相关指标: {}个".format(len(cta)))
    
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


def calculate_all_scores(observation_end_date=None, weekly=True):
    """计算所有策略评分"""
    print("\n" + "=" * 60)
    print("正在计算策略评分...")
    print("=" * 60)
    
    scorer = StrategyScorer(observation_end_date, weekly=weekly)
    all_scores = scorer.get_all_scores()
    
    for strategy_name, result in all_scores.items():
        print(f"\n{strategy_name}: {result['score']}分 - {result['level']}")
        for detail in result['details']:
            print(f"  • {detail}")
    
    return all_scores


def generate_report(all_scores, all_indicators):
    """生成控制台报告"""
    meta = all_indicators.get('_meta') or {}
    obs = meta.get('observation_end_date')
    if obs:
        print("\n观测数据截止日: {}".format(obs))
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

def run_strategy_monitor(output_excel=True, excel_path=None, observation_end_date=None, weekly=True):
    """
    策略环境监测主函数（默认周度：数据截止到「上个自然周五」对齐的交易日）
    :param output_excel: 是否输出Excel报告
    :param excel_path: Excel文件保存路径，默认自动生成
    :param observation_end_date: 数据截止日 YYYY-MM-DD；None 且 weekly=True 时自动取上周观测日
    :param weekly: True 使用周度截止日；False 用 observation_end_date 或当日
    :return: (all_scores, all_indicators, excel_path)
    """
    run_ts = dt.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    if weekly and observation_end_date is None:
        observation_end_date = get_weekly_observation_end_date()
    elif not weekly and observation_end_date is None:
        observation_end_date = dt.datetime.now().strftime('%Y-%m-%d')
    
    print("\n" + "=" * 60)
    print("策略环境监测系统（周度监测）" if weekly else "策略环境监测系统")
    print("运行时间: {}".format(run_ts))
    if observation_end_date:
        print("观测数据截止日（行情/截面）: {}".format(observation_end_date))
    print("=" * 60)
    
    # 1. 收集所有指标
    all_indicators = collect_all_indicators(observation_end_date, weekly=weekly)
    
    # 2. 计算所有策略评分
    all_scores = calculate_all_scores(observation_end_date, weekly=weekly)
    
    # 3. 生成控制台报告
    generate_report(all_scores, all_indicators)
    
    # 4. 生成Excel报告
    excel_file_path = None
    if output_excel:
        print("\n" + "=" * 60)
        print("正在生成Excel报告...")
        print("=" * 60)
        generator = ExcelReportGenerator(
            output_path=excel_path,
            observation_end_date=observation_end_date,
            run_time_str=run_ts,
            weekly=weekly)
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
