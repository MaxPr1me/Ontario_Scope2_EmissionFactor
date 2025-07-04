%%% DO NOT REDISTRIBUTE %%%
%Code written by Max St-Jacques
%Last Update 2023-09-08
%Code for 2023 Jan-Aug
%Using MATLAB 2022a

%% Upload Data

%Clear screen, variables, figures
clc;
clear all;
close all;

%Read in Generator Data

%Direct generator data as downloaded from IESO data directory
filename{1} = 'C:\Users\maxpr\Google Drive\Academics\EF CODE\Supply\GOC-2023.xlsx';
%Direct zonal demand data as downloaded from IESO data directory
filename{2} = 'C:\Users\maxpr\Google Drive\Academics\EF CODE\Demand\PUB_DemandZonal_2023.csv';
%Custom generator location and technology list
% Updated to use the CSV version of the generator list
filename{3} = 'Generator_List.csv';

%Direct power flows from other state/province connections as downloaded from IESO data directory
filename{4} = 'C:\Users\maxpr\Google Drive\Academics\EF CODE\Transmission\PUB_IntertieScheduleFlowYear_2023.csv';


for i = 1:4
    [num{i},txt{i},raw{i}] = xlsread (filename{i});
end

for i = 1:4
    num{i}(isnan(num{i}))=0;
end

%Timestep
time = datetime(txt{1}(2:end,1))+ hours(num{1}(1:end,1));


%Naming convention of Nodes
% 1. Northwest
% 2. Northeast
% 3. Ottawa
% 4. East
% 5. Toronto
% 6. Essa
% 7. Bruce
% 8. Southwest
% 9. Niagara
% 10. West

%Demand of each zone

for i = 1:10
    Demand{i} = num{2}(1:end,2+i);
end

%Naming convention of Tech
% 1. Biofuel
% 2. Hydro
% 3. Natural Gas
% 4. Nuclear
% 5. Solar
% 6. Wind

%Generators being split by zone and technology

%Create all generator combination dataset
for i = 1:10
    for j= 1:6
         Gen{i}{j}(1:size(time,1),1) = zeros(size(time,1),1);
    end
end

gencount = 0;
%Loop through generator data
for i = 4:size(raw{1},2)
    %Loop through generator list
    for j = 2: size(raw{3},1)
        %Look for same generator name
        if strcmp(raw{1}(1,i),raw{3}(j,1)) == 1     
            %Values for Region
            if strcmp(raw{3}(j,3),'Northwest') == 1
                Region = 1;
            elseif strcmp(raw{3}(j,3),'Northeast') == 1
                Region = 2;
            elseif strcmp(raw{3}(j,3),'Ottawa') == 1
                Region = 3;
            elseif strcmp(raw{3}(j,3),'East') == 1
                Region = 4;
            elseif strcmp(raw{3}(j,3),'Toronto') == 1
                Region = 5;
            elseif strcmp(raw{3}(j,3),'Essa') == 1
                Region = 6;
            elseif strcmp(raw{3}(j,3),'Bruce') == 1
                Region = 7;
            elseif strcmp(raw{3}(j,3),'Southwest') == 1
                Region = 8;
            elseif strcmp(raw{3}(j,3),'Niagara') == 1
                Region = 9;
            elseif strcmp(raw{3}(j,3),'West') == 1
                Region = 10;    
            end
              
            %Values for Technology
            if strcmp(raw{3}(j,2),'Biofuel') == 1
                Tech = 1;
            elseif strcmp(raw{3}(j,2),'Hydro') == 1
                Tech = 2;
            elseif strcmp(raw{3}(j,2),'Gas') == 1
                Tech = 3;
            elseif strcmp(raw{3}(j,2),'Nuclear') == 1
                Tech = 4;
            elseif strcmp(raw{3}(j,2),'Solar') == 1
                Tech = 5;
            elseif strcmp(raw{3}(j,2),'Wind') == 1
                Tech = 6; 
            end
                 
            %Adding Data to correct Gen{Region}{Tech}
            Gen{Region}{Tech} = [Gen{Region}{Tech} num{1}(1:end,i-1)];
            
            %Remove NaN values
            Gen{Region}{Tech}(isnan(Gen{Region}{Tech}))=0;
            
            %Break out of loop and move to next generator
            gencount = gencount+1;
           break;
        end
    end
end

%% 
%Export/Import from other Provinces and States

%Make sure they are integers
num{4}(:,2:46) = round(num{4}(:,2:46));

Manitoba = num{4}(:,4);
Michigan = num{4}(:,10);
Minnesota = num{4}(:,13);
NewYork = num{4}(:,16);
QCNE = num{4}(:,25) + num{4}(:,31);
QCOtt = num{4}(:,19) + num{4}(:,28) + num{4}(:,34) + num{4}(:,37) + num{4}(:,40);
QCE = num{4}(:,22) + num{4}(:,43);


%% First level Emission Factor for Ontario

EF = zeros(size(time,1),10);
GenOutput = zeros(size(time,1),10);

%Ontario wide EF
OntarioEF = zeros(size(time,1),1);
OntarioGenOutput = zeros(size(time,1),1);

%TechER represents emission rates for certain technologies
%TechER(2,1) represents that hydro generators emit 0 emissions
TechER = [6.15;0;525;0.15;6.15;0.74]; %in t CO2e/GWh
TechER = TechER / 1000; %in t CO2e/MWh


%Loop through each hour
for timestep = 1:size(time,1) 
    %Loop through regions
    for Region = 1:10
        %Loop through technologies
        for Tech = 1:6
            GenOutput(timestep,Region) = GenOutput(timestep,Region) + sum(Gen{Region}{Tech}(timestep,:));
            EF(timestep,Region) = EF(timestep,Region) + sum(Gen{Region}{Tech}(timestep,:)) * TechER(Tech,1);
            
            OntarioGenOutput(timestep,1) = OntarioGenOutput(timestep,1) + sum(Gen{Region}{Tech}(timestep,:));
            OntarioEF(timestep,1) = OntarioEF(timestep,1) + sum(Gen{Region}{Tech}(timestep,:)) * TechER(Tech,1);
        end
        EF(timestep,Region) = EF(timestep,Region) / GenOutput(timestep,Region);
        %Remove NaN values
        EF(isnan(EF))=0;
    end
    OntarioEF(timestep,1) = OntarioEF(timestep,1) / OntarioGenOutput(timestep,1);
end

%% Ontario EF update with grid trade

AvgOntarioEF = mean(OntarioEF);
Quebec = QCE(:,1) + QCNE(:,1) + QCOtt(:,1);

%Ontario Demand
OntarioDemand = zeros(size(time,1),1);
for i = 1:10
    OntarioDemand(:,1) = OntarioDemand(:,1) + Demand{i}(:,1);
end

%NewEFs for Ontario
NewOntarioEF = zeros(size(time,1),1);

%EFs from these regions
EFTrade = [2.2; 502; 463; 211; 1.7]; %in t CO2e/GWh
EFTrade = EFTrade / 1000; %in t CO2e/MWh

%Loop through each hour
for timestep = 1:size(time,1)
    
    %Positive number is export and negative is import
    if Manitoba(timestep,1) < 0
        Trade(1,1) = Manitoba(timestep,1)*-1;
    else
        Trade(1,1) = 0;
    end
    
    if Michigan(timestep,1) < 0
        Trade(2,1) = Michigan(timestep,1)*-1;
    else
        Trade(2,1) = 0;
    end
    
    if Minnesota(timestep,1) < 0
        Trade(3,1) = Minnesota(timestep,1)*-1;
    else
        Trade(3,1) = 0;
    end
    
    if NewYork(timestep,1) < 0
        Trade(4,1) = NewYork(timestep,1)*-1;
    else
        Trade(4,1) = 0;
    end
    
    if Quebec(timestep,1) < 0
        Trade(5,1) = Quebec(timestep,1)*-1;
    else
        Trade(5,1) = 0;
    end
    
    %New Ontario EF
    
    SelfSupplied = OntarioDemand(timestep,1) - Trade(1,1) - Trade(2,1) - Trade(3,1) - Trade(4,1) - Trade(5,1);
    for i = 1:5
        NewOntarioEF(timestep,1) = NewOntarioEF(timestep,1) + EFTrade(i,1)*Trade(i,1);
    end;
    NewOntarioEF(timestep,1) = NewOntarioEF(timestep,1)+ OntarioEF(timestep,1)*SelfSupplied;
    NewOntarioEF(timestep,1) = NewOntarioEF(timestep,1)/ OntarioDemand(timestep,1);
end

NewAvgOntarioEF = mean(NewOntarioEF);

%% Linear Program for Every Zone

%Equality Constraint
Aeq = zeros (18,50);
beq = zeros (18,1);

 
%Inequality Constrains
Aineq = zeros (12,50);
Aineq(1,1) = 1;
Aineq(2,4) = 1;
Aineq(3,5) = 1;
Aineq(4,8) = 1;
Aineq(5,12) = 1;
Aineq(6,13) = 1;
Aineq(7,14) = 1;
Aineq(8,15) = 1;
Aineq(9,16:17) = 1;
Aineq(10,18) = 1;
Aineq(11,19) = 1;
Aineq(12,21) = 1;

bineq = [325;350;2100;2900;2000;1500;2000;7500;5000;3000;1990;1800];

%Linprog values
CostVector = ones(50,1);
CostVector(31:50,1) = 1000;
CostVector([37 47],1) = 10;
LowerBound = zeros (50,1);

%Nortwest
Aeq(1, [1 2 3]) = 1;
Aeq(1, [4 23 25]) = -1;
%Norteast
Aeq(2, [4 5 6]) = 1;
Aeq(2, [1 13 28]) = -1;
%Ottawa
Aeq(3, 7) = 1;
Aeq(3, [8 29]) = -1;
%East
Aeq(4, 8:10) = 1;
Aeq(4, [11 26 30]) = -1;
%Toronto
Aeq(5, 11:12) = 1;
Aeq(5, [14 16]) = -1;
%Essa
Aeq(6, 13:14) = 1;
Aeq(6, [5 12 17]) = -1;
%Bruce
Aeq(7, 15) = 1;
%Southwest
Aeq(8, 16:18) = 1;
Aeq(8, [15 19 21]) = -1;
%Niagara
Aeq(9, 19:20) = 1;
Aeq(9, 27) = -1;
%West
Aeq(10, 21:22) = 1;
Aeq(10, [18 24]) = -1;
%Manitoba
Aeq(11, 2) = 1;
Aeq(11, 23) = -1;
%Michigan
Aeq(12, 22) = 1;
Aeq(12, 24) = -1;
%Minnesota
Aeq(13, 3) = 1;
Aeq(13, 25) = -1;
%New York
Aeq(14, [9 20]) = 1;
Aeq(14, 26:27) = -1;
%Quebec to NE
Aeq(15, 6) = 1;
Aeq(15, 28) = -1;
%Quebec to Ottawa
Aeq(16, 7) = 1;
Aeq(16, 29) = -1;
%Quebec to E
Aeq(17, 10) = 1;
Aeq(17, 30) = -1;
%Balance
Aeq(18, 31:40) = 1;
Aeq(18, 41:50) = -1;

for i = 1:10
    Aeq(i,i+30) = 1;
    Aeq(i,i+40) = -1;
end

%For everytimestep
for timestep = 1:size(time,1)
    
    for Region = 1:10
        %Equality constraint
        beq(Region,1) = GenOutput(timestep,Region) - Demand{Region}(timestep,1);
    end
    
    beq(11,1) = Manitoba(timestep,1);
    beq(12,1) = Michigan(timestep,1);
    beq(13,1) = Minnesota(timestep,1);
    beq(14,1) = NewYork(timestep,1);
    beq(15,1) = QCNE(timestep,1);
    beq(16,1) = QCOtt(timestep,1);
    beq(17,1) = QCE(timestep,1);
    
    beq(18,1) = (sum(beq(1:10,1))-sum(beq(11:17,1)));
    
    
    options = optimoptions('linprog','Algorithm','dual-simplex','Display','none','OptimalityTolerance',1.0000e-07);
    [x{timestep},fval{timestep},exitflag{timestep},output{timestep}] = linprog(CostVector,Aineq,bineq,Aeq,beq,LowerBound,[],[],options);

end

%% Total Energy Being Moved

TotalTransfer = zeros(50,1);

for timestep = 1:size(time,1)
    TotalTransfer(:,1) = TotalTransfer(:,1) + x{timestep}(:,1);
end

%% Sub-Region EF update with grid trade

AvgEF(1,:) = mean(EF(:,:));

for timestep = 1:size(time,1)
    
    P(1:10,1) = GenOutput(timestep,:)';
    P(11,1) = x{timestep}(23,1);
    P(12,1) = x{timestep}(24,1);
    P(13,1) = x{timestep}(25,1);
    P(14,1) = x{timestep}(26,1)+x{timestep}(27,1);
    P(15,1) = x{timestep}(28,1)+x{timestep}(29,1)+x{timestep}(30,1);    
%     P(16,1) = 0;
%     for i = 41:50
%         P(16,1) = P(16,1) + x{timestep}(i,1);
%     end
    
    for Region = 1:10
        D(Region,1) = Demand{Region}(timestep,1); 
    end
    D(11,1) = x{timestep}(2,1);
    D(12,1) = x{timestep}(22,1);
    D(13,1) = x{timestep}(3,1);
    D(14,1) = x{timestep}(9,1)+x{timestep}(20,1);
    D(15,1) = x{timestep}(6,1)+x{timestep}(7,1)+x{timestep}(10,1);
%     D(16,1) = 0;
%     for i = 31:40
%       D(16,1) = D(16,1) + x{timestep}(i,1);
%     end
    
    T = zeros(15,15);
    T(1,2) = x{timestep}(1,1);
    T(1,11) = x{timestep}(2,1);
    T(1,13) = x{timestep}(3,1);
    T(2,1) = x{timestep}(4,1);
    T(2,6) = x{timestep}(5,1);
    T(2,15) = x{timestep}(6,1);
    T(3,15) = x{timestep}(7,1);
    T(4,3) = x{timestep}(8,1);
    T(4,14) = x{timestep}(9,1);
    T(4,15) = x{timestep}(10,1);
    T(5,4) = x{timestep}(11,1);
    T(5,6) = x{timestep}(12,1);
    T(6,2) = x{timestep}(13,1);
    T(6,5) = x{timestep}(14,1);
    T(7,8) = x{timestep}(15,1);
    T(8,5) = x{timestep}(16,1);
    T(8,6) = x{timestep}(17,1);
    T(8,10) = x{timestep}(18,1);
    T(9,8) = x{timestep}(19,1);
    T(9,14) = x{timestep}(20,1);
    T(10,8) = x{timestep}(21,1);
    T(10,12) = x{timestep}(22,1);
    T(11,1) = x{timestep}(23,1);
    T(12,10) = x{timestep}(24,1);
    T(13,1) = x{timestep}(25,1);
    T(14,4) = x{timestep}(26,1);
    T(14,9) = x{timestep}(27,1);
    T(15,2) = x{timestep}(28,1);
    T(15,3) = x{timestep}(29,1);
    T(15,4) = x{timestep}(30,1);
%     for i = 1:10
%         T(i,16) = x{timestep}(30+i,1);
%         T(16,i) = x{timestep}(40+i,1);
%     end
 
    F(1,1:10) = EF(timestep,:);
    F(1,11:15) = [2.2 502 463 211 1.7]; %in t CO2e/GWh
    F(1,11:15) = F(1,11:15) / 1000; %in t CO2e/MWh
%     F(1,16) = 0;
 
    for i = 1:10
        X(i,1) = D(i,1) + sum(T(i,:));
    end
    
    X(11,1) =x{timestep}(2,1)+x{timestep}(23,1);
    X(12,1) = x{timestep}(22,1)+x{timestep}(24,1);
    X(13,1) =x{timestep}(3,1)+x{timestep}(25,1);
    X(14,1) = x{timestep}(9,1)+ x{timestep}(20,1)+ x{timestep}(26,1) +x{timestep}(27,1);
    X(15,1) = x{timestep}(6,1)+x{timestep}(7,1)+x{timestep}(10,1) ...
        +x{timestep}(28,1)+x{timestep}(29,1)+x{timestep}(30,1);
%     X(16,1) = sum(x{timestep}(31:40,1)) + sum(x{timestep}(41:50,1));

        
    if all(X) == 1
        NewX = inv(diag(X));
    else
        NewX = pinv(diag(X));
    end
    
    
    B=NewX*T;
    
    G=inv(eye(15)-B);
    
    H = G*diag(D)*NewX;

    Eg = F'.*P;
    
    Ec = diag(Eg)*H;
    Ec2 = ones(1,15)*Ec;
    
     if all(diag(D)) == 1
        NewD=inv(diag(D));
    else
        NewD=pinv(diag(D));
    end
    
    
    NewEF(timestep,:)=Ec2*NewD;
   
end

NewAvgEF(1,:) = mean(NewEF(:,:));


%% Print CSV

col_header={'Timestep','Ontario','Northwest','Northeast','Ottawa','East','Toronto',...
    'Essa','Bruce','Southwest','Niagara','West'};

data = OntarioEF;
data(:,2:11)= EF;
data = data*1000;

xlswrite('My_file2023.xls',data,'Supply-based EF','B2');     %Write data
xlswrite('My_file2023.xls',col_header,'Supply-based EF','A1');     %Write column header
xlswrite('My_file2023.xls',cellstr(time),'Supply-based EF','A2');      %Write row header


data2 = NewOntarioEF;
data2(:,2:11)= NewEF(:,1:10);
data2 = data2*1000;

xlswrite('My_file2023.xls',data2,'Demand-based EF','B2');     %Write data
xlswrite('My_file2023.xls',col_header,'Demand-based EF','A1');     %Write column header
xlswrite('My_file2023.xls',cellstr(time),'Demand-based EF','A2');      %Write row header


%% Next

