-- CTE, Window Functions: Determine for which Microsoft employees, their 
-- directly previous employer was Google.

with lags as (
    select
        user_id,
        employer,
        case when 
        	lag(end_date) over(
        		order by user_id, start_date
        	) = start_date then 1 else 0 end as consecutive,
        lag(employer) over(order by user_id, start_date) as prev_empl
    from linkedin_users
)
select
    count(distinct(user_id))
from lags
where consecutive = 1 and employer = 'Google' and prev_empl = 'Microsoft';