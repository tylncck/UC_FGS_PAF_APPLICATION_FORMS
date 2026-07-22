SELECT 
    UCID = DSCT.Emplid
    , Reporting_Name = DP.Person_Reporting_Name
    , Term = DT.Term_Desc
    , Degree_Plan = CONCAT(Degree_Plan_Desc, ' (', Degree_Plan, ')')
    , Academic_Plan = CONCAT(Primary_Acad_Plan_Desc, ' (', Primary_Acad_Plan, ')') 
    , Admit_Term = Admit_Term_Desc
    , Cumulative_GPA = CAST(ISNULL(GPA.Cumulative_GPA, 0) AS DECIMAL(8, 2))
    , Total_Units_Taken = CAST(ISNULL(CT.Units_Taken, 0) AS INT)
FROM 
    UofC_DW_Student.dbo.D_Student_Career_Term AS DSCT 

LEFT JOIN
    UofC_DW_Student.dbo.D_Term AS DT
ON
    DT.Term_Key = DSCT.Term_Key

LEFT JOIN   
    UofC_DW_Student.dbo.D_Person_PS AS DP
ON  
    DP.Person_Key = DSCT.Person_Key

OUTER APPLY (
    SELECT 
        Cumulative_GPA = CAST(SUM(Units_Taken * Grade_Points_Per_Unit) / SUM(Units_Taken) AS NUMERIC(8,3))
    FROM    
        UofC_DW_Student.dbo.F_Student_Courses_And_Credits AS FSCC
    WHERE
        FSCC.Academic_Career = 'GRAD'
        AND FSCC.Include_In_GPA = 'Y'
        AND FSCC.Repeat_Code IN ('', 'HIGH')
        AND FSCC.Student_Enrol_Status = 'E'
        AND FSCC.Units_Attempted = 'Y'
        AND FSCC.Term BETWEEN DSCT.Admit_Term AND DSCT.Term
        AND FSCC.Emplid = DSCT.Emplid
) AS GPA

OUTER APPLY (
    SELECT 
        Units_Taken = SUM(Units_Taken)
    FROM 
        UofC_DW_Student.dbo.F_Student_Courses_And_Credits AS FSCC
    WHERE   
        FSCC.Academic_Career = DSCT.Academic_Career
        AND FSCC.Emplid = DSCT.Emplid
        AND FSCC.Academic_Program = DSCT.Academic_Program
        AND FSCC.Include_In_GPA = 'Y'
        AND FSCC.Repeat_Code IN ('', 'HIGH')
        AND FSCC.Student_Enrol_Status = 'E'
        AND FSCC.Units_Attempted = 'Y'
) AS CT

WHERE  
    0=0
    AND DSCT.Emplid = ? --This has to be ? because with python we will inject emplids one by one while running the report for each applicant. 
    AND DSCT.Term = (
        SELECT TOP 1
            DSCT2.Term
        FROM
            UofC_DW_Student.dbo.D_Student_Career_Term AS DSCT2
        WHERE   
            DSCT2.Emplid = DSCT.Emplid
            AND DSCT2.Term <= (SELECT Current_Term FROM RO_REPORTING_STAGE.dbo.tvf_Active_Term(DEFAULT))
            AND DSCT2.Term >= 2261
            AND DSCT2.Program_Status = 'AC'
        ORDER BY 
            DSCT2.Term DESC
    )