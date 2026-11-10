# @Author: YiLong Li
# @Date: 2026/10/28 15:44
# @File: execute.py
# @Software: PyCharm


"""
备件模块单位工时求解   Python 3.10+
依赖：pip install numpy scipy openpyxl threadpoolctl
所有模式（含神经网络）均用这些依赖；不再需要安装PyTorch。
输入：data/Resource.xlsx、output/A_result.xlsx。
输出：output/execute_result.xlsx、output/execute_explain.xlsx。
退出码：0=当前任务完成；1=有模块未通过/求解未通过；2=输入或输出失败。
"""

from __future__ import annotations

# ====================== 用户配置 ======================
MODE = '2.1'

# 1.1 严格方程：线性规划检查非负可行性及严格正性，再验证原始工时等式。
# 1.2 非负最小二乘：专用线性求解器；最终值仍必须>0且>= MIN_RESULT_SECONDS 秒。
# 1.3 相对误差加权：分母使用原始总工时，不使用扣除固定值后的余额。
# 1.4 Huber稳健拟合：降低极端误差影响，不删除异常样本。
# 1.5 正则化最小二乘：在原始单位工时上加L2惩罚，参数见下方。

# 2.1 对数参数MAP：log(单位工时)服从正态先验，不是完整贝叶斯后验。
# 2.2 Bootstrap集成：各模块分别统计有效次数；缺失/无法识别不填0。

# 3.1 正权重线性神经网络：验证集早停选轮数，再用全部数据重训。
# 3.2 非线性神经网络：仅预测对照，不能产生每模块固定单位工时。
# 不自动切换模式。3.2结果表会写“不适用：非线性预测”，不要用于单位工时。

MIN_RESULT_SECONDS = 0.5    # 结果验收阈值(求解下限)(单位:s)，不是优化下限；固定正数不受此限制。

# 求解结果小于此阈值判错，等于可通过。不把0/小数值强行改为 MIN_RESULT_SECONDS 秒。
SEED = 42
CPU_THREADS = 2             # 防止BLAS线程过多拖慢计算；需要threadpoolctl。
MAX_ITER = 3000
OPTIMIZER_RESTARTS = 2      # 同一目标的数值精修次数，不切换模式/不放宽验收。
OPTIMIZER_FTOL = 1e-15      # 防止目标函数变化提前停止但梯度仍偏大。
LBFGS_MEMORY = 30
SOLVER_TOL = 1e-10
OPTIMALITY_TOL = 1e-6       # 在归一化目标上检查一阶最优性，不是业务下限。
LINEAR_INNER_MAX_ITER = 10000
STRICT_ATOL = 1e-6          # 严格方程的绝对容差，单位：小时。
STRICT_RTOL = 1e-6
STRICT_POSITIVE_TOL = 1e-9  # 线性规划缩放变量的正性确认容差。
LP_TIME_LIMIT = 120.0
RANK_RTOL = 1e-10           # 列归一化后SVD数值秩相对容差。
IDENTIFIABILITY_TOL = 1e-8  # 零空间分量阈值；不可识别模块显示红色。
RELATIVE_FLOOR_HOURS = 1.0
HUBER_HOURS = 20.0
REGULARIZATION = 0.01
REGULARIZATION_REFERENCE_HOURS = 0.01
# 1.5目标：mean(((Ax-y)/工时尺度)^2)/2 + lambda*sum((x/参考小时)^2)/2。
PRIOR_HOURS = 0.01        # 2.1对数正态先验中位数，小时/件。
PRIOR_LOG_STD = 2.0
NOISE_HOURS = 30.0
BOOTSTRAP_RUNS = 50
BOOTSTRAP_MIN_VALID_FRACTION = 0.5
BOOTSTRAP_MIN_VALID_COUNT = 10
# 每次抽样都检查列覆盖及可识别性；区间为条件经验分位数，不保证覆盖率。
ENABLE_VALIDATION = True
VALIDATION_FRACTION = 0.2
MIN_VALIDATION_SAMPLES = 20
# 随机留出验证：矩阵无日期，不能称为未来时间验证。训练未覆盖的模块不填0。
NETWORK_EPOCHS = 3000
NETWORK_LR = 0.01
NETWORK_HIDDEN = 32
NETWORK_PATIENCE = 100
NETWORK_MIN_IMPROVEMENT = 1e-6
# 3.x必须启用验证；达到训练上限且未早停，不当作训练已稳定。
# 固定值=绿色，可独立确定=蓝色，估计值=黄色，错误/*/不适用=红色。
# 结果C列保留小时/件；D列新增秒/件=C列×3600。错误候选值只在解释报告展示。
# ======================================================

import hashlib
import math
import shutil
import sys
import tempfile
import time
from contextlib import nullcontext
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Callable
from uuid import uuid4

try:
    import numpy as np
    import scipy
    from scipy.linalg import svd
    from scipy.optimize import linprog, lsq_linear, minimize
    from scipy.sparse import csr_matrix, diags, eye, hstack, vstack
    from scipy.special import expit
    import openpyxl
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
except ImportError as error:
    raise SystemExit(f'缺少依赖：{error}。请运行：pip install numpy scipy openpyxl threadpoolctl')

BASE = Path(__file__).resolve().parent
RESOURCE = BASE / 'data' / 'Resource.xlsx'
INPUT = BASE / 'output' / 'A_result.xlsx'
OUT = BASE / 'output'
GREEN, BLUE, RED = 'C6EFCE', 'BDD7EE', 'FFC7CE'
MODE_NAMES = {'1.1':'严格方程', '1.2':'最小二乘', '1.3':'相对误差加权',
              '1.4':'Huber稳健', '1.5':'正则化', '2.1':'对数参数MAP',
              '2.2':'Bootstrap集成', '3.1':'正权重线性网络', '3.2':'非线性预测'}
CONFIG_KEYS = [k for k,v in list(globals().items()) if k.isupper() and
               isinstance(v,(str,bool,int,float)) and k not in ('GREEN','BLUE','RED')]


@dataclass
class Inputs:
    modules: list[tuple[str,str]]
    fixed: np.ndarray
    A: np.ndarray
    b: np.ndarray
    lines: list[int]
    sources: list[list]


@dataclass
class Fit:
    x: np.ndarray
    ok: bool
    message: str
    iterations: int = 0
    optimality: float | None = None
    per_ok: np.ndarray | None = None
    counts: np.ndarray | None = None
    low: np.ndarray | None = None
    high: np.ndarray | None = None
    trace: list[list] = field(default_factory=list)
    predict_function: Callable | None = None
    epochs: int = 0

    def predict(self, X):
        if self.predict_function is not None:
            return np.asarray(self.predict_function(X), dtype=float)
        missing = ~np.isfinite(self.x)
        result = np.asarray(X @ np.nan_to_num(self.x, nan=0., posinf=0., neginf=0.))
        if missing.any():
            result[np.any(X[:,missing] != 0, axis=1)] = np.nan
        return result


class InputError(ValueError):
    def __init__(self, issues, modules=None, fixed=None):
        self.issues, self.modules, self.fixed = issues, modules or [], fixed
        super().__init__(f'预检查发现{len(issues)}项错误')


def log(message):
    print(message, flush=True)


def finite_positive(value):
    return not isinstance(value, bool) and isinstance(value,(int,float)) and math.isfinite(value) and value > 0


def validate_config():
    errors=[]
    if MODE not in MODE_NAMES: errors.append(['配置','MODE','','未知求解模式',str(MODE)])
    positives=('MIN_RESULT_SECONDS','OPTIMIZER_FTOL','SOLVER_TOL','OPTIMALITY_TOL','STRICT_ATOL','STRICT_RTOL',
               'STRICT_POSITIVE_TOL','LP_TIME_LIMIT','RANK_RTOL','IDENTIFIABILITY_TOL',
               'RELATIVE_FLOOR_HOURS','HUBER_HOURS','REGULARIZATION_REFERENCE_HOURS',
               'PRIOR_HOURS','PRIOR_LOG_STD','NOISE_HOURS','NETWORK_LR','NETWORK_MIN_IMPROVEMENT')
    for key in positives:
        if not finite_positive(globals()[key]): errors.append(['配置',key,'','必须是有限正数',str(globals()[key])])
    for key in ('CPU_THREADS','MAX_ITER','OPTIMIZER_RESTARTS','LBFGS_MEMORY','LINEAR_INNER_MAX_ITER','BOOTSTRAP_RUNS',
                'BOOTSTRAP_MIN_VALID_COUNT','MIN_VALIDATION_SAMPLES','NETWORK_EPOCHS',
                'NETWORK_HIDDEN','NETWORK_PATIENCE'):
        v=globals()[key]
        if isinstance(v,bool) or not isinstance(v,int) or v<=0: errors.append(['配置',key,'','必须是正整数',str(v)])
    if not isinstance(SEED,int) or isinstance(SEED,bool) or SEED<0: errors.append(['配置','SEED','','必须是非负整数',str(SEED)])
    if not isinstance(ENABLE_VALIDATION,bool): errors.append(['配置','ENABLE_VALIDATION','','必须是True或False',str(ENABLE_VALIDATION)])
    for key in ('VALIDATION_FRACTION','BOOTSTRAP_MIN_VALID_FRACTION'):
        v=globals()[key]
        if not finite_positive(v) or v>=1: errors.append(['配置',key,'','必须大于0且小于1',str(v)])
    if not isinstance(REGULARIZATION,(int,float)) or isinstance(REGULARIZATION,bool) or not math.isfinite(REGULARIZATION) or REGULARIZATION<0:
        errors.append(['配置','REGULARIZATION','','必须是非负有限数',str(REGULARIZATION)])
    if isinstance(MIN_VALIDATION_SAMPLES,int) and MIN_VALIDATION_SAMPLES<2:
        errors.append(['配置','MIN_VALIDATION_SAMPLES','','至少为2',''])
    if not errors and (BOOTSTRAP_RUNS<2 or BOOTSTRAP_MIN_VALID_COUNT<2 or BOOTSTRAP_MIN_VALID_COUNT>BOOTSTRAP_RUNS):
        errors.append(['配置','BOOTSTRAP','','有效次数应>=2且不超过总次数',''])
    for key in ('RANK_RTOL','IDENTIFIABILITY_TOL'):
        if finite_positive(globals()[key]) and globals()[key]>=1:
            errors.append(['配置',key,'','必须小于1',str(globals()[key])])
    if MODE in ('3.1','3.2') and not ENABLE_VALIDATION:
        errors.append(['配置','ENABLE_VALIDATION','','神经网络模式必须启用验证',''])
    if errors: raise InputError(errors)


def fingerprint(path):
    h=hashlib.sha256()
    with path.open('rb') as stream:
        for block in iter(lambda:stream.read(1024*1024),b''): h.update(block)
    return h.hexdigest()


def read_inputs():
    errors=[]; modules=[]; fixed=[]; sources=[]; before={}
    for path in (RESOURCE,INPUT):
        if not path.is_file(): errors.append(['文件',str(path),'','文件不存在',''])
        else: before[path]=(path.stat().st_size,path.stat().st_mtime_ns)
    if errors: raise InputError(errors)
    w=load_workbook(RESOURCE,read_only=True,data_only=False)
    try:
        if '模块信息' not in w.sheetnames: raise ValueError('缺少“模块信息”表单')
        it=iter(w['模块信息'].values); header=next(it,())
        headers=[str(v).strip() if v is not None else '' for v in header]
        if any(headers.count(n)!=1 for n in ('备件模块','自变量')): raise ValueError('备件模块、自变量字段缺失或重复')
        names=[n for n in ('值','确定值') if n in headers]
        if len(names)!=1 or headers.count(names[0])!=1: raise ValueError('必须且只能有一个“值”或“确定值”字段')
        nc,vc,fc=(headers.index(n) for n in ('备件模块','自变量',names[0]))
        seen=set()
        for line,r in enumerate(it,2):
            if all(v is None for v in r): continue
            name,var,value=(r[c] if c<len(r) else None for c in (nc,vc,fc))
            name='' if name is None else str(name).strip();var='' if var is None else str(var).strip()
            if not name or name in seen: errors.append(['模块信息','备件模块',line,'名称为空或重复',name])
            if var!=f'x{len(modules)+1}': errors.append(['模块信息','自变量',line,'必须按x1、x2…顺序排列',var])
            seen.add(name);modules.append((name,var))
            try:
                if value is None or (isinstance(value,str) and not value.strip()): number=np.nan
                else:
                    if isinstance(value,bool): raise ValueError()
                    number=float(value)
                    if not math.isfinite(number) or number<=0: raise ValueError()
                fixed.append(number)
            except (ValueError,TypeError,OverflowError):
                fixed.append(np.nan);errors.append(['模块信息',names[0],line,'固定值必须为有限正数或空白',str(value)])
        if not modules: errors.append(['模块信息','','','没有模块数据',''])
    except (ValueError,KeyError) as e: errors.append(['Resource','','',str(e),''])
    finally: w.close()
    if errors: raise InputError(errors,modules,np.asarray(fixed))
    w=load_workbook(INPUT,read_only=True,data_only=False); data=[];lines=[]
    try:
        sheet=w.worksheets[0];sheet.reset_dimensions()
        if sheet.title!='汇总矩阵': errors.append(['A_result','表单','','第一张表应为汇总矩阵',sheet.title])
        for line,r in enumerate(sheet.values,1):
            if all(v is None for v in r): continue
            if len(r)!=len(modules)+1:
                errors.append(['A_result','列数',line,f'应为{len(modules)+1}列',str(len(r))]);continue
            bad=[i+1 for i,v in enumerate(r) if isinstance(v,bool) or not isinstance(v,(int,float)) or not math.isfinite(v) or v<0]
            if bad:
                errors.append(['A_result','数值',line,'含空白、文本、公式、负数或非有限数','列:'+','.join(map(str,bad))]);continue
            data.append(r);lines.append(line)
    finally: w.close()
    if not data: errors.append(['A_result','','','没有有效数据',''])
    for path in (RESOURCE,INPUT):
        if before[path]!=(path.stat().st_size,path.stat().st_mtime_ns):errors.append(['文件',str(path),'','读取过程中被修改',''])
        sources.append([str(path.resolve()),before[path][0],fingerprint(path)])
    if errors: raise InputError(errors,modules,np.asarray(fixed))
    a=np.asarray(data,dtype=float)
    return Inputs(modules,np.asarray(fixed),a[:,:-1],a[:,-1],lines,sources)


def diagnose(X):
    if not X.shape[1]: return np.zeros(0,dtype=bool),0,None
    scales=np.linalg.norm(X,axis=0)
    if not np.all(np.isfinite(scales)) or np.any(scales<=0): raise ValueError('诊断矩阵包含全零列或溢出')
    _,s,vh=svd(X/scales,full_matrices=False,check_finite=False)
    rank=int(np.count_nonzero(s>s[0]*RANK_RTOL))
    ambiguous=np.zeros(X.shape[1],dtype=bool)
    if rank<X.shape[1]: ambiguous=np.maximum(0.,1.-np.sum(vh[:rank]**2,axis=0))>IDENTIFIABILITY_TOL
    condition=float(s[0]/s[rank-1]) if rank else None
    return ambiguous,rank,condition


def failed_fit(n,message,x=None):
    return Fit(np.full(n,np.nan) if x is None else x,False,message)


def minimize_refined(objective,initial,bounds=None):
    """保留目标函数，严格检查最终梯度；有剩余预算时从当前解继续精修。"""
    current=initial.copy();iterations=0;best=None
    for attempt in range(OPTIMIZER_RESTARTS+1):
        remaining=MAX_ITER-iterations
        if remaining<=0:break
        r=minimize(objective,current,jac=True,method='L-BFGS-B',bounds=bounds,
                   options={'maxiter':remaining,'ftol':OPTIMIZER_FTOL,
                            'gtol':min(SOLVER_TOL,OPTIMALITY_TOL/10),
                            'maxls':50,'maxcor':LBFGS_MEMORY})
        iterations+=int(r.nit)
        loss,gradient=objective(r.x)
        stationarity=gradient if bounds is None else r.x-np.maximum(0.,r.x-gradient)
        r['optimality']=float(np.max(np.abs(stationarity)))
        r['checked_ok']=bool(r.success and np.isfinite(loss) and np.all(np.isfinite(r.x))
                             and np.isfinite(r.optimality) and r.optimality<=OPTIMALITY_TOL)
        if best is None or (np.isfinite(loss) and (not np.isfinite(best.fun) or loss<=best.fun)):
            best=r
        if r.checked_ok:best=r;break
        if not np.all(np.isfinite(r.x)):break
        current=r.x.copy()
    best['nit']=iterations
    return best


def linear_fit(X,y,mode,total):
    n=X.shape[1]
    if not n:return Fit(np.zeros(0),True,'没有待求模块')
    scales=np.linalg.norm(X,axis=0)
    if not np.all(np.isfinite(scales)) or np.any(scales<=0):return failed_fit(n,'无效列尺度')
    Z=csr_matrix(X/scales)
    ys=max(float(np.sqrt(np.mean(y*y))),1.)
    target=y/ys
    if mode=='1.1':
        eq=hstack([Z,csr_matrix((len(y),1))],format='csr')
        constraints=hstack([-eye(n),csr_matrix(np.ones((n,1)))],format='csr')
        r=linprog(np.r_[np.zeros(n),-1.],A_ub=constraints,b_ub=np.zeros(n),
                  A_eq=eq,b_eq=target,bounds=[(0,None)]*n+[(0,1)],method='highs',
                  options={'time_limit':LP_TIME_LIMIT,'primal_feasibility_tolerance':1e-9})
        x=np.full(n,np.nan) if r.x is None else r.x[:-1]*ys/scales
        ok=bool(r.success and r.x[-1]>STRICT_POSITIVE_TOL and np.all(x>0))
        if r.status==2: message='非负等式在求解器容差下不可行'
        elif r.success and not ok: message='仅找到边界解，或严格正性在数值容差内无法确认'
        else: message=str(r.message)
        return Fit(x,ok,message,int(r.nit or 0))
    if mode=='1.4':
        # Huber是线性预测上的凸损失。用NNLS初值和解析梯度避免通用非线性迭代停滞。
        warm=linear_fit(X,y,'1.2',total)
        if np.all(np.isfinite(warm.x)):initial=np.maximum(0.,warm.x*scales/ys)
        else:initial=np.full(n,1.)
        delta=HUBER_HOURS/ys
        def objective(z):
            residual=Z@z-target;absolute=np.abs(residual)
            loss=np.where(absolute<=delta,.5*residual**2,delta*(absolute-.5*delta)).sum()
            return float(loss),np.asarray(Z.T@np.clip(residual,-delta,delta))
        r=minimize_refined(objective,initial,[(0,None)]*n)
        x=r.x*ys/scales
        return Fit(x,bool(r.checked_ok),f'Huber解析梯度精修：{r.message}；投影梯度={r.optimality:.6g}',
                   int(r.nit)+warm.iterations,float(r.optimality))
    M=Z;target2=target
    if mode=='1.3':
        weights=ys/np.maximum(np.abs(total),RELATIVE_FLOOR_HOURS)
        M=diags(weights)@Z;target2=target*weights
    elif mode=='1.5' and REGULARIZATION:
        penalty=np.sqrt(len(y)*REGULARIZATION)*ys/(scales*REGULARIZATION_REFERENCE_HOURS)
        M=vstack([Z,diags(penalty)],format='csr');target2=np.r_[target,np.zeros(n)]
    # 加权/正则项会重新改变列尺度；二次列平衡只改变坐标，不改变原目标函数。
    balance=np.sqrt(np.asarray(M.power(2).sum(axis=0)).ravel())
    if not np.all(np.isfinite(balance)&(balance>0)):return failed_fit(n,'加权矩阵列尺度无效')
    balanced=M@diags(1./balance)
    r=lsq_linear(balanced,target2,bounds=(0,np.inf),method='trf',lsq_solver='lsmr',
                 tol=SOLVER_TOL,lsmr_tol=1e-12,max_iter=MAX_ITER,lsmr_maxiter=LINEAR_INNER_MAX_ITER)
    x=r.x/balance*ys/scales
    ok=bool(r.success and np.all(np.isfinite(x)) and float(r.optimality)<=OPTIMALITY_TOL)
    message=f'{r.message}；列平衡后归一化最优性指标={r.optimality:.6g}'
    if r.success and not ok:message+='；未通过独立停止条件验收'
    return Fit(x,ok,message,int(r.nit or 0),float(r.optimality))


def map_fit(X,y):
    # 在t=log(x)空间的MAP：不声称这是x空间的众数，也不声称是后验均值。
    scale=max(len(y),1);matrix=csr_matrix(X)
    def objective(t):
        with np.errstate(over='ignore',invalid='ignore'):
            x=np.exp(t);res=(matrix@x-y)/NOISE_HOURS;u=(t-np.log(PRIOR_HOURS))/PRIOR_LOG_STD
            loss=(np.dot(res,res)+np.dot(u,u))/(2*scale)
            grad=(x*(matrix.T@res)/NOISE_HOURS+u/PRIOR_LOG_STD)/scale
        if not np.isfinite(loss) or not np.all(np.isfinite(grad)):
            return np.inf,np.full_like(t,1e100)
        return float(loss),grad
    r=minimize_refined(objective,np.full(X.shape[1],np.log(PRIOR_HOURS)))
    with np.errstate(over='ignore'):x=np.exp(r.x)
    ok=bool(r.checked_ok and np.all(np.isfinite(x)&(x>0)))
    return Fit(x,ok,f'log(x)空间MAP精修：{r.message}；平均目标梯度={r.optimality:.6g}',int(r.nit),float(r.optimality))


def bootstrap_fit(X,y,total,progress=True):
    baseline=fit_model(X,y,'1.2',total,False)
    rng=np.random.default_rng(SEED);draws=np.full((BOOTSTRAP_RUNS,X.shape[1]),np.nan);trace=[]
    for k in range(BOOTSTRAP_RUNS):
        ids=rng.integers(0,len(y),len(y));sample=X[ids];covered=np.any(sample!=0,axis=0)
        if covered.any():
            result=fit_model(sample[:,covered],y[ids],'1.2',total[ids],False)
            accepted=np.zeros(covered.sum(),dtype=bool)
            if result.ok:
                try:
                    ambiguous,_,_=diagnose(sample[:,covered]);accepted=~ambiguous & np.isfinite(result.x)
                    draws[k,np.flatnonzero(covered)[accepted]]=result.x[accepted]
                except Exception as e:
                    result.ok=False;result.message=f'本次抽样诊断失败：{e}'
            trace.append([k+1,int(covered.sum()),int(accepted.sum()),result.ok,result.message])
        else:trace.append([k+1,0,0,False,'没有覆盖待求模块'])
        if progress and ((k+1)%max(1,BOOTSTRAP_RUNS//10)==0 or k==0):log(f'[Bootstrap] {k+1}/{BOOTSTRAP_RUNS}')
    counts=np.sum(np.isfinite(draws),axis=0);needed=max(BOOTSTRAP_MIN_VALID_COUNT,math.ceil(BOOTSTRAP_RUNS*BOOTSTRAP_MIN_VALID_FRACTION))
    sufficient=counts>=needed;x=baseline.x.copy();low=np.full(X.shape[1],np.nan);high=low.copy()
    for j,count in enumerate(counts):
        if count:
            vals=draws[np.isfinite(draws[:,j]),j];x[j]=np.median(vals)
            if sufficient[j]:low[j],high[j]=np.quantile(vals,[.025,.975])
    # 覆盖不足的候选值可保留，但绝不充当已接受模块值；不以0补缺失。
    return Fit(x,bool(sufficient.any()),f'各模块需至少{needed}次有效抽样；区间为覆盖且可识别条件下的经验分位数',
               BOOTSTRAP_RUNS,per_ok=sufficient,counts=counts,low=low,high=high,trace=trace)


def fit_model(X,y,mode,total,progress=True):
    try:
        if not X.shape[1]:return Fit(np.zeros(0),True,'没有待求模块')
        if mode in ('1.1','1.2','1.3','1.4','1.5'):result=linear_fit(X,y,mode,total)
        elif mode=='2.1':result=map_fit(X,y)
        elif mode=='2.2':result=bootstrap_fit(X,y,total,progress)
        else:return failed_fit(X.shape[1],'神经网络需通过验证流程调用')
        if mode=='1.1':
            difference=result.predict(X)-y
            if not np.all(np.isfinite(difference)) or not np.all(np.abs(difference)<=STRICT_ATOL+STRICT_RTOL*np.abs(total)):
                result.ok=False;result.message+='；原始小时等式验收未通过'
        return result
    except Exception as e:return failed_fit(X.shape[1],f'{type(e).__name__}: {e}')


def neural_fit(X,y,mode,Xv=None,yv=None,epochs=None,progress=True):
    """NumPy实现全批量Adam；验证早停或按已选轮数重训。"""
    n=X.shape[1];rng=np.random.default_rng(SEED)
    scale=np.sqrt(np.mean(X*X,axis=0));ys=max(float(np.sqrt(np.mean(y*y))),1.)
    if np.any(scale<=0):return failed_fit(n,'神经网络训练列不能全零')
    Z=csr_matrix(X/scale);target=y/ys;nonlinear=mode=='3.2'
    gate=np.any(X!=0,axis=1).astype(float)
    if nonlinear:
        params=[rng.normal(0,1/np.sqrt(n),(n,NETWORK_HIDDEN)),np.zeros(NETWORK_HIDDEN),
                rng.normal(0,1/np.sqrt(NETWORK_HIDDEN),NETWORK_HIDDEN),np.zeros(1)]
    else:
        initial=max(float(np.mean(np.maximum(y,0)))/max(float(np.mean(X.sum(axis=1))),1.),1e-6)
        beta=initial*scale/ys
        params=[beta+np.log(-np.expm1(-beta))]
    def forward(z,g,p):
        if not nonlinear:return z@np.logaddexp(0,p[0]),None
        hidden=np.tanh(z@p[0]+p[1]);raw=hidden@p[2]+p[3][0]
        return g*np.logaddexp(0,raw),(hidden,raw)
    def loss_grad(p):
        pred,cache=forward(Z,gate,p);res=pred-target;deriv=2*res/len(y)
        if not nonlinear:grad=[(Z.T@deriv)*expit(p[0])]
        else:
            hidden,raw=cache;d=deriv*gate*expit(raw);dh=(d[:,None]*p[2])*(1-hidden*hidden)
            grad=[Z.T@dh,dh.sum(axis=0),hidden.T@d,np.array([d.sum()])]
        return float(np.mean(res*res)),grad
    use_validation=epochs is None
    if use_validation and (Xv is None or not len(Xv)):return failed_fit(n,'没有可用验证样本，无法执行早停')
    zv=None if Xv is None else csr_matrix(Xv/scale);gv=None if Xv is None else np.any(Xv!=0,axis=1).astype(float)
    m=[np.zeros_like(p) for p in params];v=[np.zeros_like(p) for p in params]
    best=[p.copy() for p in params];best_score=np.inf;best_epoch=0;anchor=np.inf;bad=0;stopped=False;trace=[]
    budget=NETWORK_EPOCHS if epochs is None else epochs
    try:
        for epoch in range(1,budget+1):
            loss,grads=loss_grad(params)
            if not np.isfinite(loss) or not all(np.all(np.isfinite(g)) for g in grads):raise FloatingPointError('网络损失或梯度非有限数')
            norm=np.sqrt(sum(np.sum(g*g) for g in grads));factor=min(1.,10./max(norm,1e-100))
            for j,g in enumerate(grads):
                g=g*factor;m[j]=.9*m[j]+.1*g;v[j]=.999*v[j]+.001*g*g
                params[j]-=NETWORK_LR*(m[j]/(1-.9**epoch))/(np.sqrt(v[j]/(1-.999**epoch))+1e-8)
            train_score=float(np.mean((forward(Z,gate,params)[0]-target)**2))
            score=float(np.mean((forward(zv,gv,params)[0]-yv/ys)**2)) if use_validation else train_score
            if not np.isfinite(score):raise FloatingPointError('网络验证损失非有限数')
            if score<best_score:best_score=score;best=[p.copy() for p in params];best_epoch=epoch
            if score<anchor-NETWORK_MIN_IMPROVEMENT:anchor=score;bad=0
            else:bad+=1
            if epoch==1 or epoch%100==0:trace.append([epoch,train_score,score if use_validation else None])
            if progress and epoch%500==0:log(f'[网络] {epoch}/{budget}；归一化损失{score:.6g}')
            if use_validation and bad>=NETWORK_PATIENCE:stopped=True;break
        # 全量重训按选定轮数输出最后参数，不以训练误差再选择另一轮数。
        chosen=best if use_validation else params
        def predict(new):return forward(csr_matrix(new/scale),np.any(new!=0,axis=1).astype(float),chosen)[0]*ys
        x=np.full(n,np.nan) if nonlinear else np.logaddexp(0,chosen[0])*ys/scale
        ok=stopped if use_validation else True
        text=('验证早停' if stopped else '达到训练上限，尚未确认稳定') if use_validation else '按验证选定轮数全量重训完成'
        trace.append([epoch,loss_grad(chosen)[0],best_score if use_validation else None])
        return Fit(x,ok,text,epoch,trace=trace,predict_function=predict,epochs=best_epoch if use_validation else budget)
    except Exception as e:
        chosen=best
        def predict(new):return forward(csr_matrix(new/scale),np.any(new!=0,axis=1).astype(float),chosen)[0]*ys
        x=np.full(n,np.nan) if nonlinear else np.logaddexp(0,chosen[0])*ys/scale
        return Fit(x,False,f'网络失败：{e}',trace=trace,predict_function=predict,epochs=best_epoch)


def validate_and_fit(X,y,total):
    val_predictions=np.full(len(y),np.nan);val_notes=['未参与留出验证']*len(y);validation=[];val_fit=None
    is_network=MODE in ('3.1','3.2')
    if ENABLE_VALIDATION and len(y)>=MIN_VALIDATION_SAMPLES and X.shape[1]:
        rng=np.random.default_rng(SEED);order=rng.permutation(len(y))
        count=max(1,min(len(y)-1,round(len(y)*VALIDATION_FRACTION)));vid=order[:count];tid=order[count:]
        covered=np.any(X[tid]!=0,axis=0)
        eligible=~np.any(X[vid][:,~covered]!=0,axis=1)
        for i,can in zip(vid,eligible):val_notes[i]='验证候选预测' if can else '训练集未覆盖该行部分模块，未计入误差'
        log(f'[验证] 训练{len(tid)}行；留出{len(vid)}行；可评估{int(eligible.sum())}行')
        if covered.any():
            if is_network:
                val_fit=neural_fit(X[tid][:,covered],y[tid],MODE,X[vid[eligible]][:,covered],y[vid[eligible]])
            else:val_fit=fit_model(X[tid][:,covered],y[tid],MODE,total[tid])
            if eligible.any():val_predictions[vid[eligible]]=val_fit.predict(X[vid[eligible]][:,covered])
        else:
            val_fit=Fit(np.zeros(0),True,'训练集没有未知模块业务');val_predictions[vid[eligible]]=0.
        validation=[['验证划分','固定随机种子留出；无日期，非未来时间验证'],['验证训练行数',len(tid)],
                    ['验证留出行数',len(vid)],['验证可覆盖行数',int(eligible.sum())],
                    ['验证求解状态',val_fit.ok],['验证说明',val_fit.message]]
        comparable=vid[np.isfinite(val_predictions[vid])]
        validation += [['训练均值基准 '+k,v] for k,v in metrics(total[comparable],np.full(len(comparable),float(np.mean(total[tid]))))]
        validation += [['基准说明','只用训练集总工时均值预测同一批可评估行；低于此基准的误差才表示优于常数预测']]
        if not val_fit.ok:
            for i in vid[eligible]:val_notes[i]='验证求解未通过；以下仅为候选预测'
    else:
        validation=[['验证说明','已关闭、样本不足或没有待求模块；未报告泛化能力']]
    log(f'[求解] 使用全部{len(y)}个样本执行模式{MODE}')
    if is_network and X.shape[1]:
        if val_fit is None or not val_fit.epochs:
            fit=failed_fit(X.shape[1],'缺少有效验证训练，不能选择神经网络轮数')
        elif not val_fit.ok:
            candidate=np.full(X.shape[1],np.nan);candidate[covered]=val_fit.x
            def candidate_predict(new):
                p=val_fit.predict(new[:,covered]);p[np.any(new[:,~covered]!=0,axis=1)]=np.nan
                return p
            fit=Fit(candidate,False,'验证未通过，未执行全量重训；保留验证阶段候选：'+val_fit.message,
                    val_fit.iterations,trace=[['验证训练']+r for r in val_fit.trace],
                    predict_function=candidate_predict,epochs=val_fit.epochs)
        else:
            fit=neural_fit(X,y,MODE,epochs=val_fit.epochs)
            fit.ok=fit.ok and val_fit.ok
            fit.message+=f'；选定{val_fit.epochs}轮；验证阶段：{val_fit.message}'
            fit.trace=[['验证训练']+r for r in val_fit.trace]+[['全量重训']+r for r in fit.trace]
    else:fit=fit_model(X,y,MODE,total)
    return fit,val_predictions,val_notes,validation


def classify(fixed,active,ambiguous,fit):
    values=fixed.copy();values[active]=fit.x
    accepted=np.isfinite(fixed);texts=[];reasons=[];local=0
    for j in range(len(fixed)):
        why=[]
        if np.isfinite(fixed[j]):text=float(fixed[j]);why=['固定值，未调整']
        elif not active[j]:text='*';why=['无业务样本']
        else:
            x=values[j]
            if not fit.ok:why.append('求解未通过')
            if MODE=='3.2':why.append('不适用：非线性预测')
            else:
                if fit.per_ok is not None and not fit.per_ok[local]:why.append('有效重采样不足')
                if not np.isfinite(x):why.append('无有效候选值')
                elif x<=0:why.append('求解错误：非正数')
                elif x*3600<MIN_RESULT_SECONDS:why.append(f'求解错误：小于{MIN_RESULT_SECONDS:g}秒')
            accepted[j]=not why
            text=float(x) if not why else '；'.join(why)
            local+=1
        texts.append(text)
        reasons.append('；'.join(why) if why else
                       '估计值，需复核：无法独立确定' if ambiguous[j] else
                       '可独立确定：通过数值检查，仍需业务核实')
    return values,accepted,texts,reasons


def metrics(actual,pred):
    mask=np.isfinite(pred)
    if not mask.any():return [['可计算误差行数',0]]
    e=pred[mask]-actual[mask];b=actual[mask];positive=b>0
    rows=[['可计算误差行数',int(mask.sum())],['RMSE(小时)',float(np.sqrt(np.mean(e*e)))],
          ['MAE(小时)',float(np.mean(np.abs(e)))],['平均偏差(小时)',float(np.mean(e))],['最大绝对误差(小时)',float(np.max(np.abs(e)))]]
    if positive.any():rows.append(['MAPE(排除实际0工时)',float(np.mean(np.abs(e[positive]/b[positive])))])
    denom=float(np.sum((b-np.mean(b))**2))
    rows.append(['R²',None if denom==0 else 1-float(e@e)/denom])
    return rows


def seconds_cell(value):
    """正式结果仅换算已验收/固定数值；错误候选值留在解释报告。"""
    if isinstance(value,(int,float,np.number)) and not isinstance(value,bool):
        seconds=float(value)*3600
        return seconds if math.isfinite(seconds) else '换算溢出'
    if value=='*': return '*'
    return '不适用' if '不适用' in str(value) else '未通过'


def acceptance_counts(values,active,ambiguous,fit):
    if MODE=='3.2':return [['数值求解器是否通过',fit.ok],['单位工时验收','不适用：本模式仅预测，无固定模块工时']]
    candidate=values[active]
    return [
        ['数值求解器是否通过',fit.ok],
        ['无有效候选值',int(np.sum(~np.isfinite(candidate)))],
        ['非正候选值',int(np.sum(np.isfinite(candidate)&(candidate<=0)))],
        [f'低于{MIN_RESULT_SECONDS:g}秒的正候选值',int(np.sum((candidate>0)&(candidate*3600<MIN_RESULT_SECONDS)))],
        ['无法独立确定模块',int(np.sum(ambiguous))],
        ['有效重采样不足',int(np.sum(~fit.per_ok)) if fit.per_ok is not None else 0],
    ]


def build_reports(data,run_id):
    known=np.isfinite(data.fixed);active=np.any(data.A!=0,axis=0)&~known
    fixed_hours=data.A[:,known]@data.fixed[known];y=data.b-fixed_hours;X=data.A[:,active]
    hours_unique=int(np.unique(data.b).size)
    if hours_unique==1:
        log(f'[数据提示] 全部{len(data.b)}行总工时均为{data.b[0]:g}小时；需核实是否为出勤工时，不能据此保证模块工时可靠')
    if not np.all(np.isfinite(y)):raise ValueError('固定工时计算溢出')
    log(f'[预检查] 样本{len(y)}；模块{len(known)}；固定{int(known.sum())}；待求{int(active.sum())}')
    log('[诊断] 检查业务矩阵的数值秩和模块可识别性')
    local_ambiguous,rank,condition=diagnose(X);ambiguous=np.zeros(len(known),dtype=bool);ambiguous[active]=local_ambiguous
    fit,vpred,vnotes,validation=validate_and_fit(X,y,data.b)
    values,accepted,texts,reasons=classify(data.fixed,active,ambiguous,fit)
    prediction=fixed_hours+fit.predict(X)
    # 严格模式中无待求模块时也必须检查等式；固定值仍始终原样保留。
    if MODE=='1.1' and not np.all(np.isfinite(prediction)&(np.abs(prediction-data.b)<=STRICT_ATOL+STRICT_RTOL*np.abs(data.b))):
        fit.ok=False;fit.message+='；全矩阵严格等式未通过'
        values,accepted,texts,reasons=classify(data.fixed,active,ambiguous,fit)
    rejected=active&~accepted
    estimated=active&accepted&ambiguous
    identified=active&accepted&~ambiguous
    status='求解未通过' if not fit.ok else '部分模块未通过验收' if rejected.any() else '试算完成，含待复核估计值' if estimated.any() else '完成'
    if MODE=='3.2' and fit.ok:status='预测对照完成；不提供固定单位工时'
    complete=bool(fit.ok and not rejected.any() and MODE!='3.2')
    log(f'[数值求解] {"通过" if fit.ok else "未通过"}；模块验收另行判断')
    log(f'[结果] 可独立确定{int(identified.sum())}；估计值待复核{int(estimated.sum())}；未通过{int(rejected.sum())}；无业务{int((~active&~known).sum())}')
    if estimated.any():log('[说明] 估计值可用于试算；总量拟合良好不代表每个模块工时准确。')
    if not fit.ok:log(f'[求解状态] {fit.message}')
    result=[['备件模块','自变量','求解值（小时/件）','换算值（秒/件）','结果状态']]+[[n,v,texts[i],seconds_cell(texts[i]),reasons[i]] for i,(n,v) in enumerate(data.modules)]
    diagnostics=[['备件模块','自变量','结果状态','出现样本数','固定值(小时/件)','候选值(小时/件)',
                  '候选值(秒/件)','有效重采样次数','经验下界(小时/件)','经验上界(小时/件)','补充说明']]
    idx=0
    for j,(name,var) in enumerate(data.modules):
        count=lo=hi=None
        if active[j]:
            if fit.counts is not None:count=int(fit.counts[idx]);lo=fit.low[idx];hi=fit.high[idx]
            idx+=1
        note='固定值低于验收阈值，按规则原样保留' if known[j] and data.fixed[j]*3600<MIN_RESULT_SECONDS else ''
        if known[j] and not np.any(data.A[:,j]):note+='；本期无业务，保留固定值'
        diagnostics.append([name,var,reasons[j],int(np.count_nonzero(data.A[:,j])),data.fixed[j],values[j],values[j]*3600,count,lo,hi,note])
    samples=[['A_result原始行号','实际工时(小时)','固定工时(小时)','候选预测工时(小时)',
              '候选误差(预测-实际)','候选相对误差','输出数值能否试算本行','输出数值试算工时(小时)',
              '留出验证候选工时(小时)','留出验证候选误差','验证说明','异常说明']]
    computable=~np.any(data.A[:,~accepted]!=0,axis=1)
    if not fit.ok or MODE=='3.2':computable[:]=False
    verified=data.A@np.where(accepted,np.nan_to_num(values,nan=0.),0.)
    val_total=fixed_hours+vpred
    for i,line in enumerate(data.lines):
        p=prediction[i];e=p-data.b[i];warnings=[]
        if fixed_hours[i]>data.b[i]+STRICT_ATOL:warnings.append('固定工时超过实际，未截断/未删除')
        if not np.any(data.A[i]) and data.b[i]>0:warnings.append('模块全零但总工时大于0')
        if np.any(data.A[i]) and data.b[i]==0:warnings.append('有业务但总工时为0')
        if np.any(data.A[i,estimated]!=0):warnings.append('使用待复核估计值，试算工时不代表模块实测结果')
        if not computable[i]:warnings.append('含未接受模块或任务未通过，候选误差不代表可用结果')
        samples.append([line,data.b[i],fixed_hours[i],p,e,e/data.b[i] if data.b[i]>0 else None,
                        '是' if computable[i] else '否',verified[i] if computable[i] else None,
                        val_total[i],val_total[i]-data.b[i],vnotes[i],'；'.join(warnings)])
    summary=[['项目','内容'],['运行标识',run_id],['模式',MODE+' '+MODE_NAMES[MODE]],['状态',status],
             ['求解器验收',fit.ok],['求解说明',fit.message],['迭代次数',fit.iterations],['最优性指标',fit.optimality],
             ['当前有业务模块是否均有可试算数值',complete],['当前有业务模块是否均可独立确定',bool(complete and not estimated.any())],['结果单位','C列小时/件；D列=C列×3600，单位秒/件；错误候选值仅在解释报告展示'],
             ['验收下限(秒)',MIN_RESULT_SECONDS],['下限用途','仅验收未知模块；不参与优化，不修改固定值'],
             ['总工时不同取值数',hours_unique],['总工时最小值(小时)',float(np.min(data.b))],
             ['总工时最大值(小时)',float(np.max(data.b))],
             ['工时口径提示','全部工时相同：请核实出勤工时与实际业务作业工时；常数预测可获得零误差，不能据此验证单位工时' if hours_unique==1 else '请确保工时与业务量属于同一天、同一业务范围'],
             ['样本数',len(y)],['模块数',len(known)],['固定模块数',int(known.sum())],['待求模块数',int(active.sum())],
             ['可独立确定模块数',int(identified.sum())],['估计值待复核模块数',int(estimated.sum())],['未通过待求模块数',int(rejected.sum())],
             ['无业务未知模块(*)',int((~active&~known).sum())],['无法独立确定数',int(ambiguous.sum())],
             ['数值秩',rank],['非零奇异值条件数',condition],
             ['可识别性说明','仅依据当前业务矩阵与数值容差；正则/先验不能替代缺失信息，不保证业务准确'],
             ['固定工时超过实际行数',int(np.sum(fixed_hours>data.b+STRICT_ATOL))],
             ['误差口径','完整候选解含未通过模块；验证也是候选预测，不能冒充已验收结果'],
             ['验证限制','随机留出不等于未来验证；网络用同一验证集选轮数，不能作为独立测试集']]
    summary += [['原因统计（可重叠） '+k,v] for k,v in acceptance_counts(values,active,ambiguous,fit)]
    summary += [['全量候选 '+k,v] for k,v in metrics(data.b,prediction)]
    summary += [['输出数值试算 '+k,v] for k,v in metrics(data.b,np.where(computable,verified,np.nan))]
    summary += validation + [['留出候选 '+k,v] for k,v in metrics(data.b,val_total)]
    if MODE=='2.1':
        summary += [['先验中位数(秒/件)',PRIOR_HOURS*3600],['MAP说明','正数估计受配置先验影响；通过0.5秒门槛不等于实际测得的标准工时']]
    if MODE=='1.5':
        summary += [['正则化说明','L2惩罚会向0收缩；强度和参考尺度应结合验证误差及业务信息选择，不能为提高通过率随意调参']]
    summary += [['运行配置','详见配置参数表单'],
                ['依赖版本',f'numpy={np.__version__}; scipy={scipy.__version__}; openpyxl={openpyxl.__version__}']]
    detail=[['类型','位置','行号','说明','原始值']]
    if hours_unique==1:detail.append(['数据提示','总工时','','全部样本总工时相同；需核实实际作业工时口径',float(data.b[0])])
    if not fit.ok:detail.append(['求解','全量','','求解未通过',fit.message])
    for j in np.flatnonzero(estimated):detail.append(['待复核估计',data.modules[j][1],'',reasons[j],values[j]])
    for j in np.flatnonzero(rejected):detail.append(['模块',data.modules[j][1],'',reasons[j],values[j]])
    for i in np.flatnonzero(fixed_hours>data.b+STRICT_ATOL):detail.append(['样本','A_result',data.lines[i],'固定工时超过实际工时',float(fixed_hours[i]-data.b[i])])
    help_rows=[['现象','处理建议'],
               ['数值求解未通过','查看求解说明和最优性指标；严格方程不可行时不能靠增加迭代次数解决，可选择近似模式评估'],
               ['低于0.5秒','核对数量单位和实际作业工时；通过现场测时补充固定值，不强行抬高候选值'],
               ['无法独立确定','数值已作为待复核估计输出；补充独立样本或实测固定值后重算，不能仅凭拟合误差确认单模块工时'],
               ['总工时始终相同','核实是否使用了固定出勤工时；如包含等待、辅助工作或非同日业务，应补充对应记录，而不是凭空调整工时'],
               ['模式2.1通过更多模块','结果受先验影响，需以独立测时核实；通过率不代表准确率'],
               ['模式3.2没有模块取值','该模式仅用于非线性预测，不提供固定单位工时'],
               ['部分模块未通过','数值求解完成与模块验收分开；详见原因统计及模块诊断，退出码1并不一定代表程序崩溃']]
    sheets=[('求解总览',summary),('样本拟合',samples),('模块诊断',diagnostics),('异常记录',detail),
            ('输入来源',[['文件','字节数','SHA256']]+data.sources),
            ('配置参数',[['参数','本次取值']]+[[k,globals()[k]] for k in CONFIG_KEYS]),('问题处理',help_rows)]
    if fit.trace:
        if MODE=='2.2':header=['次数','覆盖模块数','可识别有效模块数','求解器通过','说明']
        else:header=['阶段','轮数','训练归一化损失','验证归一化损失']
        sheets.append(('过程记录',[header]+fit.trace))
    code=0 if (complete or (MODE=='3.2' and fit.ok)) else 1
    return [('模块信息',result)],sheets,known,code,status


def clean_cell(v):
    if isinstance(v,np.generic):v=v.item()
    if isinstance(v,float) and not math.isfinite(v):return None
    if isinstance(v,str) and len(v)>32000:return v[:31980]+'（内容过长已截断）'
    return v


def write_book(path,sheets,known,run_id):
    w=Workbook();w.remove(w.active);w.properties.description=f'execute run_id={run_id}'
    try:
        for title,rows in sheets:
            s=w.create_sheet(title);s.freeze_panes='A2';s.sheet_view.showGridLines=False
            for row_index,row in enumerate(rows,1):
                s.append([clean_cell(v) for v in row])
                for cell in next(s.iter_rows(min_row=row_index,max_row=row_index,max_col=len(row))):
                    if isinstance(cell.value,str):cell.data_type='s'
                    cell.font=Font(name='微软雅黑',size=10)
                    cell.alignment=Alignment(vertical='center',wrap_text=True)
                    if isinstance(cell.value,float):cell.number_format='0.000000E+00' if (0<abs(cell.value)<1e-4 or abs(cell.value)>=1e9) else '0.000000'
            for c in s[1]:c.fill=PatternFill('solid',fgColor='1F4E78');c.font=Font(name='微软雅黑',bold=True,color='FFFFFF');c.alignment=Alignment(vertical='center',wrap_text=True)
            s.row_dimensions[1].height=32;s.auto_filter.ref=s.dimensions
            for col in range(1,s.max_column+1):s.column_dimensions[get_column_letter(col)].width=24
            if title=='模块信息':
                s.column_dimensions['A'].width=36;s.column_dimensions['B'].width=14;s.column_dimensions['C'].width=52;s.column_dimensions['D'].width=24;s.column_dimensions['E'].width=60
                for j,row in enumerate(s.iter_rows(min_row=2)):
                    fixed=j<len(known) and known[j];cell=row[2]
                    is_estimate=len(row)>4 and '估计值' in str(row[4].value)
                    color=GREEN if fixed else 'FFF2CC' if is_estimate else BLUE if isinstance(cell.value,(int,float)) else RED
                    cell.fill=PatternFill('solid',fgColor=color)
                    cell.number_format='0.000000E+00' if isinstance(cell.value,(int,float)) and 0<abs(cell.value)<1e-9 else '0.000000000'
                    if len(row)>3:
                        row[3].fill=PatternFill('solid',fgColor=color);row[3].number_format='0.000000'
                    if len(row)>4:row[4].fill=PatternFill('solid',fgColor=color)
                    s.row_dimensions[j+2].height=32
            elif title=='求解总览':
                s.column_dimensions['A'].width=36;s.column_dimensions['B'].width=110
                for row in s.iter_rows(min_row=2):
                    if 'MAPE' in str(row[0].value):row[1].number_format='0.00%'
                    length=len(str(row[1].value or ''));s.row_dimensions[row[0].row].height=min(210,max(24,math.ceil(length/65)*18))
            elif title=='模块诊断':
                s.column_dimensions['A'].width=36;s.column_dimensions['C'].width=55;s.column_dimensions['K'].width=48
                for row in s.iter_rows(min_row=2):
                    for i in (4,5,8,9):row[i].number_format='0.000000E+00'
                    if '估计值' in str(row[2].value):row[2].fill=PatternFill('solid',fgColor='FFF2CC')
                    elif any(k in str(row[2].value) for k in ('错误','未通过','无法','不足','无业务','不适用')):row[2].fill=PatternFill('solid',fgColor=RED)
            elif title=='样本拟合':
                s.column_dimensions['K'].width=48;s.column_dimensions['L'].width=65
                for row in s.iter_rows(min_row=2):row[5].number_format='0.00%'
            elif title=='异常记录':s.column_dimensions['D'].width=70;s.column_dimensions['E'].width=65
            elif title=='问题处理':s.column_dimensions['A'].width=35;s.column_dimensions['B'].width=100
            elif title=='配置参数':s.column_dimensions['A'].width=50;s.column_dimensions['B'].width=30
            elif title=='输入来源':s.column_dimensions['A'].width=100;s.column_dimensions['C'].width=75
        w.save(path)
    finally:w.close()


def publish(result_sheets,explain_sheets,known,run_id):
    """两份文件均写好再发布；普通异常回滚。不承诺断电/强制杀进程的原子性。"""
    OUT.mkdir(parents=True,exist_ok=True);stage=Path(tempfile.mkdtemp(prefix='.execute_',dir=OUT))
    preserve=False
    targets=[OUT/'execute_result.xlsx',OUT/'execute_explain.xlsx']
    try:
        for i,sheets in enumerate((result_sheets,explain_sheets)):write_book(stage/f'new_{i}.xlsx',sheets,known,run_id)
        preserve=True
        try:
            for i,target in enumerate(targets):
                if target.exists():
                    backup=stage/f'backup_{i}.xlsx';target.replace(backup)
            for i,target in enumerate(targets):(stage/f'new_{i}.xlsx').replace(target)
            preserve=False
        except BaseException as error:
            # 包括Ctrl+C；依据磁盘状态回滚，避免中断发生在rename和记录之间。
            problems=[]
            for i in reversed(range(len(targets))):
                if not (stage/f'new_{i}.xlsx').exists():
                    try:targets[i].unlink(missing_ok=True)
                    except OSError as e:problems.append(str(e))
            for i,target in enumerate(targets):
                backup=stage/f'backup_{i}.xlsx'
                if backup.exists():
                    try:backup.replace(target)
                    except OSError as e:problems.append(str(e))
            if problems:
                preserve=True;raise RuntimeError(f'发布失败且回滚不完整；备份位于{stage}；{problems}') from error
            preserve=False
            raise RuntimeError(f'发布失败，已恢复旧文件；请关闭Excel占用：{error}') from error
    finally:
        if not preserve:shutil.rmtree(stage,ignore_errors=True)


def main():
    start=time.monotonic();run_id=datetime.now().strftime('%Y%m%d_%H%M%S')+'_'+uuid4().hex[:8]
    log(f'[开始] 模式{MODE}；固定值不动；未知模块小于{MIN_RESULT_SECONDS}秒判错')
    data=None
    try:
        validate_config();data=read_inputs()
        try:
            from threadpoolctl import threadpool_limits
            context=threadpool_limits(limits=CPU_THREADS)
        except ImportError:
            context=nullcontext();log('[提示] 未安装threadpoolctl，无法限制计算线程数')
        with context:result,explain,known,code,status=build_reports(data,run_id)
    except Exception as e:
        issues=e.issues if isinstance(e,InputError) else [['运行','','',type(e).__name__,str(e)]]
        modules=e.modules if isinstance(e,InputError) else data.modules if data is not None else []
        fixed=e.fixed if isinstance(e,InputError) else data.fixed if data is not None else None
        known=np.isfinite(fixed) if fixed is not None else np.zeros(len(modules),dtype=bool)
        rows=[['备件模块','自变量','求解值（小时/件）','换算值（秒/件）','结果状态']]
        for i,(name,var) in enumerate(modules):
            value=float(fixed[i]) if known[i] else '运行失败：详见解释报告'
            rows.append([name,var,value,seconds_cell(value),'固定值，未调整' if known[i] else '运行失败'])
        if not modules:rows.append(['无法读取模块','','运行失败：详见解释报告','未通过','运行失败'])
        result=[('模块信息',rows)];explain=[('求解总览',[['项目','内容'],['运行标识',run_id],['状态','输入/运行失败'],['说明',str(e)]]),
                                        ('异常记录',[['类型','位置','行号','说明','原始值']]+issues)]
        code=2;status='输入/运行失败';log(f'[失败] {e}；详细记录将写入解释报告')
    try:publish(result,explain,known,run_id)
    except Exception as e:log(f'[输出失败] {e}；旧文件不能作为本次结果');return 2
    log(f'[完成] {status}；用时{time.monotonic()-start:.1f}秒；输出目录：{OUT}')
    return code


if __name__=='__main__':sys.exit(main())
