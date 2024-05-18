import numpy as np
import operator as op

# Grid size
interval = 0.5
W = 70 # Width of grid
D = 25 # Depth of grid 
A = W*D
Hole = 5  # Number of holes
typenumber = 4  # Number of soil types

# Initializing grid with zeros
group_number = np.zeros((A),dtype=int)

# Load geo_matrix from CSV
# Saperate by caommas, skip first row(title)
geo_matrix = np.loadtxt('test.csv' , delimiter = "," , skiprows = 1)

# Define hole positions from previous one
Hole1 = 1    
Hole2 = 28*2   
Hole3 = 46*2   
Hole4 = 58*2    
Hole5 = 70*2    

# Assign group numbers to grid cell
# Define size of grid into two group
# Vertical
for i in range(1,74,1):
    group_number[i-1] = 1
for i in range(74,140,1):
    group_number[i-1] = 2

# Assign soil types to hole positions
# Horizontal
for j in range(1,D+1,1):
    group_number[(Hole1-1)+(j-1)*W] = geo_matrix[j-1][0]
    group_number[(Hole2-1)+(j-1)*W] = geo_matrix[j-1][1]
    group_number[(Hole3-1)+(j-1)*W] = geo_matrix[j-1][2]
    group_number[(Hole4-1)+(j-1)*W] = geo_matrix[j-1][3]
    group_number[(Hole5-1)+(j-1)*W] = geo_matrix[j-1][4]

# Calculate transition probabilities for vertical layers
# Create a zero matrix with length of get_matrix
T_t_V = np.zeros(len(geo_matrix))
soiltype_V = {}

# Count occurrences of soil types
# Size:
for i in range(np.size(geo_matrix,1)):
    for j in range(len(geo_matrix)):
        T_t_V[j] = geo_matrix[j][i]
    for k in T_t_V[0:len(T_t_V)]:
        soiltype_V[k] = soiltype_V.get(k, 0) + 1

# Sort soil types
soiltype_V = sorted(soiltype_V.items(), key=op.itemgetter(0), reverse=False)

# Initialize transition matrix for vertical layers
VPCM = np.zeros([len(soiltype_V), len(soiltype_V)])
Tmatrix_V = np.zeros([len(soiltype_V), len(soiltype_V)])

# Fill transition matrix for vertical layers
for i in range(np.size(geo_matrix,1)):
    for j in range(len(geo_matrix)):
        T_t_V[j] = geo_matrix[j][i]
    for k in range(len(T_t_V) - 1):
        for m in range(len(soiltype_V)):
            for n in range(len(soiltype_V)):
                if T_t_V[k] == soiltype_V[m][0] and T_t_V[k + 1] == soiltype_V[n][0]:
                    VPCM[m][n] += 1
                    Tmatrix_V[m][n] += 1

# Normalize transition matrix for vertical layers
count_V = np.sum(Tmatrix_V,axis=1)
for i in range(np.size(Tmatrix_V,1)):      
    for j in range(np.size(Tmatrix_V,1)):
        Tmatrix_V[i][j] = Tmatrix_V[i][j]/count_V[i]

# Define K value
K = 9.3

# Initialize horizontal transition matrix and matrix for horizontal layers
HPCM = np.zeros([len(count_V), len(count_V)])
Tmatrix_H = np.zeros([len(count_V), len(count_V)])

# Fill horizontal transition matrix and matrix for horizontal layers
for i in range(np.size(Tmatrix_H,1)):
    for j in range(np.size(Tmatrix_H,1)):
        if i == j:
            HPCM[i][j] = K*VPCM[i][j]
            Tmatrix_H[i][j] = K*VPCM[i][j]
        else:
            HPCM[i][j] = VPCM[i][j]
            Tmatrix_H[i][j] = VPCM[i][j]

# Normalize horizontal transition matrix
count_H = np.sum(Tmatrix_H,axis=1)
for i in range(np.size(Tmatrix_H,1)):      
    for j in range(np.size(Tmatrix_H,1)):
        Tmatrix_H[i][j] = Tmatrix_H[i][j]/count_H[i]

# Initialize states and current matrix
L_state = 0
M_state = 0
Q_state = 0
G_state = 0
Nx = 0
a = 0
current_matrix = np.array([[0.0,0.0,0.0,0.0]])
transitionName = np.array([[1,2,3,4]])

# Iterate through layers and grid cells to calculate new group numbers
for layer in range(2,D+1,1):
    for i in range(1,W+1,1):
        L_state = 0
        M_state = 0
        Q_state = 0
        if i > Hole1 and i < Hole2: 
            L_state = group_number[(i-2)+(layer-1)*W]-1
            M_state = group_number[(i-1)+(layer-2)*W]-1
            Q_state = group_number[(Hole2-1)+(layer-1)*W]-1
            Nx = Hole2
        elif i > Hole2 and i < Hole3:
            L_state = group_number[(i-2)+(layer-1)*W]-1
            M_state = group_number[(i-1)+(layer-2)*W]-1
            Q_state = group_number[(Hole3-1)+(layer-1)*W]-1
            Nx = Hole3
        elif i > Hole3 and i < Hole4:
            L_state = group_number[(i-2)+(layer-1)*W]-1
            M_state = group_number[(i-1)+(layer-2)*W]-1
            Q_state = group_number[(Hole4-1)+(layer-1)*W]-1
            Nx = Hole4
        elif i > Hole4 and i < Hole5:
            L_state = group_number[(i-2)+(layer-1)*W]-1
            M_state = group_number[(i-1)+(layer-2)*W]-1
            Q_state = group_number[(Hole5-1)+(layer-1)*W]-1
            Nx = Hole5

        if i == Hole1 or i == Hole2 or i == Hole3 or i == Hole4 or i == Hole5:
            a = a+1
        else:
            TV = Tmatrix_V
            TH = Tmatrix_H
            Nx_TH = Tmatrix_H
            f_sum = 0
            k_sum = 0
            for N in range(1,Nx-i,1):
                Nx_TH = np.dot(Nx_TH,Tmatrix_H)
            for f in range(0,typenumber,1):
                f_item1 = Tmatrix_H[L_state][f]
                f_item2 = Nx_TH[f][Q_state]
                f_item3 = Tmatrix_V[M_state][f] 
                f_sum = f_sum + (f_item1*f_item2*f_item3)
            for k in range(0,typenumber,1):
                k_item1 = Tmatrix_H[L_state][k]
                k_item2 = Nx_TH[k][Q_state]
                k_item3 = Tmatrix_V[M_state][k]
                k_sum = k_item1*k_item2*k_item3
                current_matrix [0][k] = k_sum/f_sum
            group_number[(i-1)+(layer-1)*W] = np.random.choice(transitionName[0],replace=True,p=current_matrix[0])


