# Writing to an excel 
# sheet using Python 
import xlwt 
from xlwt import Workbook 

# Workbook is created 
wb = Workbook() 

# add_sheet is used to create sheet. 
sheet1 = wb.add_sheet('Sheet 1')

#Number of total particles (input) - Ntotal
# Lx,Ly and Lz
#amp_v_y - amplitude of vibratory particles in Y direction
# Delta x , delta y and delta z (inter-particle distance)
#define roughness of the particles (Roughness factor) Fr .. Place them in up and down position


Nx = input ("Number of particles in x direction")
Nx = int(Nx)
Ny = input ("Number of particles in y direction")
Ny = int(Ny)
Nz = input ("Number of particles inin z direction")
Nz = int (Nz)
dp = input ("diameter of particle")
dp = int (dp)
deltax = input ("interpaticle distance in x-direction")
deltax = int (deltax)
deltay = input ("interpaticle distance in y-direction")
deltay = int (deltay)
deltaz = input ("interpaticle distance in z-direction")
deltaz = int (deltaz)
amp = input ("Amplitude of the vibration")
amp = int(amp)
Nvx = input ("Number of vibratory particles in x-direction")
Nvx = int (Nvx)
Nvz = input ("Number of vibratory particles in z direction")
Nvz = int (Nvz)
dpv = input ("Diameter of the vibrating particles")
dpv = int (dpv)
s = input ("Spacing between the vibration particles")
s = int (s)

Lx = (Nx - 1)*deltax + (dp/2)
Ly = (Ny - 1)*deltay + (dp/2)
Lz = (Ny - 1)*deltay + (dp/2)

print (Lx, Ly, Lz)

Vf = ((4/3)*(22/7)*(dp/2)**3)/(deltax**3)
print ('Volume fraction of the arrangement is', Vf)

#Nx = int(Nx)
#Ny = int(Ny)
#Nz = int(Nz)

#Nvx = (Lx/dpv)
#Nvy = (Ly/dpv)
#Nvz = (Lz/dpv)

#Nvx = int (Nvx)
#Nvy = int (Nvy)
#Nvz = int (Nvz)

#print ('Number of particles in x,y and z', Nx, Ny, Nz)
#print ('Number of wall particles in x and z', Nvx and Nvz)

# filling column 1 ie serial number
row_num=0
for i in range (1,Nx*Ny*Nz +2* (Nvx*Nvz+1)):
    k = i
    sheet1.write(row_num,1,k)
    row_num+=1

# filling column 2 ie atom type 1   
row_num=0
for j in range (1,Nx*Ny*Nz +1):
    sheet1.write(row_num,2,1)
    row_num+=1

# filling column 2 ie atom type 2   
row_num = Nz*Ny*Nx
for j in range (1,Nvx*Nvz+1):
    sheet1.write(row_num,2,2)
    row_num+=1

# filling column 2 ie atom type 3
row_num = Nz*Ny*Nx + Nvx * Nvz
for j in range (1,Nvx*Nvz+1):
    sheet1.write(row_num,2,3)
    row_num+=1

# filling column 3 ie 1

row_num=0
for j in range (1,Nx*Ny*Nz +Nvx*Nvz+1):
    sheet1.write(row_num,3,1)
    row_num+=1

# filling column 4 ie 1

row_num=0
for j in range (1,Nx*Ny*Nz +Nvx*Nvz+1):
    sheet1.write(row_num,4,1)
    row_num+=1

# filling column 5 i.e x - axis of main particles 
row_num=0
for i in range (Nz):
    for i in range (Ny):
        for i in range (1,Nx+1):
            k = (dp/2)+ deltax *(i-1)
            sheet1.write(row_num,5,k)
            row_num+=1
            
# filling column 5 i.e. x - axis of bottom vibratory particles            
row_num = Nz*Ny*Nx
for i in range (Nvz):
    for i in range (1,Nvx+1):
        n = (dpv/2)+s*(i-1)
        sheet1.write(row_num,5,n)
        row_num+=1
        
# filling column 5 i.e. x - axis of bottom vibratory particles
row_num = Nz*Ny*Nx +Nvx*Nvz

for i in range (Nvz):
    for i in range (1,Nvx+1):
        n = (dpv/2)+s*(i-1)
        sheet1.write(row_num,5,n)
        row_num+=1

# y - axis of main particles     
row_num=0
for i in range (Nz):
    for i in range (1,Ny+1):
        for q in range (Nx):
            k = (dp/2)+ deltay *(i-1)+amp/2
            sheet1.write(row_num,6,k)
            row_num+=1
            
# y - axis of lower vibratory particles            
row_num = Nz*Ny*Nx
for i in range (1,Nvx*Nvz+1):
    n = 0
    sheet1.write(row_num,6,n)
    row_num+=1

# y - axis of upper vibratory particles            
row_num = Nz*Ny*Nx +Nvx * Nvz

for i in range (1,Nvx*Nvz+1):
    n = Ly + dp/2 + amp/2
    sheet1.write(row_num,6,n)
    row_num+=1
    
# z - axis of main particles         
row_num=0
for i in range (1,Nz+1):
    for q in range (Nx*Ny):
        k = (dpv/2)+ deltaz *(i-1)+amp/2
        sheet1.write(row_num,7,k)
        row_num+=1

# z - axis of lower vibratory particles          
row_num = Nz*Ny*Nx
for i in range (0,Nvz):
    for r in range (Nvx):
        n = (dpv/2)+s*(i)
        sheet1.write(row_num,7,n)
        row_num+=1

# z - axis of upper vibratory particles          
row_num = Nz*Ny*Nx + Nvx * Nvz

for i in range (0,Nvz):
    for r in range (Nvx):
        n = (dpv/2)+s*(i)
        sheet1.write(row_num,7,n)
        row_num+=1
        
# saving file to excel
wb.save('xlwt example.xls') 



