records = LOAD 'processed_data.csv' USING PigStorage(',') AS
(center,student,age,gender, ethnic, country, item,item_difficulty,changes_frequency,effect_index, change_interval );
test_recs = GROUP records BY student;
result_test = FOREACH test_recs {
	----filter for the first step
        country_filt = FILTER country BY item == 'XXXXX_1' ;
        -- Since we only want unique country, we need to project it out
        country_proj = FOREACH country_filt GENERATE item ;
        -- Now we can find all unique country for a given filter
        country_dist = DISTINCT country_proj ;

	----filter for the first step
        country_filt_2 = FILTER country BY item == 'XXXXX_2' ;
        -- Since we only want unique country, we need to project it out
        country_proj_2 = FOREACH country_filt_2 GENERATE item ;
        -- Now we can find all unique country for a given filter
        country_dist_2 = DISTINCT country_proj_2 ;

	---middle steps omited
    GENERATE group_country AS student, COUNT(country_dist) AS result_1,
                               COUNT(country_dist_2) AS result_2 ;


	----do the same filter for different centers
        item_filt = FILTER item BY item == 'asian' ;
        -- Since we only want unique item, we need to project it out
        item_proj = FOREACH item_filter GENERATE COUNT(item) ;

	item_filt = FILTER item BY item == 'n_america' ;
        -- Since we only want unique item, we need to project it out
        item_proj = FOREACH item_filter GENERATE COUNT(item) ;

	--omitted middle parts

	--filter for different country, item, item_difficuty 

	----futher calculate
	GENERATE group_1,
    (SUM(c_1_student) / DISTINCT(sum_student) AS country_1_stu+percentage,
    (SUM(c_2_student) / DISTINCT(sum_student) AS country_2_stu+percentage,
    (SUM(i_1_student) / DISTINCT(sum_student) AS item_1_stu+percentage,
    (SUM(i_2_student) / DISTINCT(sum_student) AS item_2_stu+percentage,    
    (SUM(m_1_student) / DISTINCT(sum_student) AS male_stu+percentage,
    (SUM(f_2_student) / DISTINCT(sum_student) AS female_stu+percentage,
    AVG(c_1_result) AS country_1_result_avg,
AVG(c_2_result) AS country_1_result_avg,
AVG(f_1_result) AS female_1_result_avg,
AVG(m_1_result) AS male_1_result_avg,
AVG(ethnic_1_result) AS country_1_result_avg.
}



STORE result_test INTO /user/root/result_test;
