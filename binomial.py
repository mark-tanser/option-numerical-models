"""

PYTHON code for Binomial Pricing Model
======================================

GENERAL ADDITIVE BINOMIAL TREE MODEL (GABTM)
 
 >input is a dictionary / keyword-list containing:

 {
 'option_type': 'vanilla'|'binary'|'touch'
 'spot_price': numeric
 'strike_price': numeric
 'time_to_expiry': numeric
 'domestic_rate': numeric
 'foreign_rate': numeric
 'log_normal_volatility': numeric
 'time_steps': numeric
 }

 >output is an array of {call.value,put.value} or for 'touch' type: {onetouch.value,notouch.value}
 approximation for the process: dS=(rd - rf).S dt + sigma.S dz
 where dz describes a Weiner process random walk.

 Assuming S is distributed log-normally as x = ln(S), then:
 dx = nu dt + sigma dz
 where: nu = r - 1/2.sigma^2

 Approximating this with by discrete time steps of dt = t / N
 and defining up and down changes in x by:
 x + dxu , x - dxd, where dxu = -dxd = dx
 
 and associated probabilties for up and down movements of pu and pd, where pu + pd = 1

 Equating mean and variance gives:
 
 E[dx] = pu.dx - pd.dx = nu.dt
 E[dx^2] = pu.dx^2 + pd.dx^2 = sigma^2.dt + nu^2.dt^2
 
 solving simultaneously gives:
  
 dx = sqrt( sigma^2.dt + nu^2.dt )
 pu = 1/2 + 1/2 . nu.dt / dx
 and we know that pd = 1 -pu
 
 with the spot price S on node(i,j) computed as:
 St(i,j) = exp(x(i,j)) = exp(x + j.dx - (i-j).dx)
 
 terminal option values V are given on the final j nodes as:
 vanilla_call = V(j,1) = Max( 0, St(j) - K )
 vanilla_put = V(j,2) = Max( 0, K - St(j) )
 call = V(j,1) = 1 if St(j) >= K ; 0 otherwise
 put = V(j,2) = 1 if St(j) < K ; 0 otherwise

"""

# import Numerical Python (NumPy) library
# to install from commandline run: 'pip install numpy'
import numpy as np 
# import standard math library
import math

def get_inputs():
	option_inputs = {}	
	option_inputs['option_type'] = 'blank'
	while not check_option_type(option_inputs['option_type']):
		option_inputs['option_type'] = raw_input('Option Type ( vanilla | binary | touch ) ?: ').lower()
	#
	option_inputs['spot_price'] = float(raw_input('Spot Price ?: '))
	option_inputs['strike_price'] = float(raw_input('Strike Price ?: '))
	option_inputs['time_to_expiry'] = float(raw_input('Time to Expiry ( years ) ?: '))
	option_inputs['domestic_rate'] = float(raw_input('Domestic Rate of Return ?: '))
	option_inputs['foreign_rate'] = float(raw_input('Foreign Rate of Return ?: '))
	option_inputs['log_normal_volatility'] = float(raw_input('Log-Normal Volatility (annualised) ?: '))
	option_inputs['time_steps'] = int(raw_input('Number of Time Steps ?: '))
	print option_inputs
	return option_inputs

def check_option_type(input_type):
	if not( input_type == 'vanilla' or input_type == 'binary' or input_type == 'touch') :
		if input_type != 'blank':
			print('Invalid option type. Please input again...')
		return False
	else:
		return True

def GABTM(option):
	OptionType = option['option_type']
	S = option['spot_price']
	K = option['strike_price']
	t = option['time_to_expiry']
	rd = option['domestic_rate']
	rf = option['foreign_rate']
	sigma = option['log_normal_volatility']
	N = option['time_steps']
	inputs_list = [S,K,t,rd,rf,sigma,N]
	#
	#compute coefficients and constants
	dt = t / N
	nu = rd - rf - 0.5 * (sigma ** 2)
	dxu = math.sqrt((sigma ** 2) * dt + ((nu * dt) ** 2))
	dxd = -dxu
	pu = 0.5 + (0.5 * (nu * dt / dxu))
	pd = 1 - pu
	disc = math.exp(-rd * dt)
    #
	#initialise asset prices and option values at maturity 
	V = np.zeros((N + 1,2))
	St = np.zeros(N + 1)
	St[0] = S * math.exp(N * dxd)
	for j in range(0, N + 1, 1):
		if j > 0:
			St[j] = St[j - 1] * math.exp(dxu - dxd)
		if OptionType == 'vanilla':
			V[j,0] = max(0, St[j] - K)
			V[j,1] = max(0, K - St[j])
		else:
				# binary and touch options
				if St[j] >= K:
					V[j,0] = 1
					V[j,1] = 0
				else:
					V[j,0] = 0
					V[j,1] = 1
	#
	#step back through tree to compute option values through to time zero
	for i in range(N - 1, -1, -1):
		for j in range(0, i + 1, 1):
			if j > 0:
				St[j] = St[j - 1] * math.exp(dxu - dxd)
			else:
				St[0] = S * math.exp(i * dxd)
			if OptionType == 'touch': #for american style options
				if St[j] >= K: #touch up
					V[j,0] = 1
				else:
					V[j,0] = disc * (pu * V[j + 1,0] + pd * V[j,0])
				#
				if St[j] <= K: #touch down
					V[j,1] = 1
				else:
					V[j,1] = disc * (pu * V[j + 1,1] + pd * V[j,1])
			else: #for european style vanilla and binary
				V[j,0] = disc * (pu * V[j + 1,0] + pd * V[j,0])
				V[j,1] = disc * (pu * V[j + 1,1] + pd * V[j,1])	
	#
	#build return price array in temporary variable 'temp'
	temp = np.zeros(2)
	if OptionType == 'touch':
		# choose touch up or touch down depending on direction of barrier from spot price
		if K >= S:
			temp[0] = V[0, 0]
		else:
			temp[0] = V[0, 1]
		temp[1] = disc - temp[0]
	else:
		temp[0] = V[0, 0]
		temp[1] = V[0, 1]
	#
	return temp


my_option = get_inputs()
option_price = GABTM(my_option)
if my_option['option_type'] == 'touch':
	print 'prices: one-touch, no-touch'
else:
	print 'prices: call, put'
print option_price

















