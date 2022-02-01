-- Joins and Subqueries: Find the customer with the highest daily total order 
-- cost between 2019-02-01 to 2019-05-01.

with daily_totals as (
    select
        cust_id,
        date,
        sum(cost) as cost
    from (
        select 
            cust_id,
            order_date as date,
            total_order_cost as cost
        from orders
        where order_date between '2019-02-01' and '2019-05-01'    
    ) as costs
    group by costs.cust_id, date
)
select c.first_name, d.date, d.cost
from daily_totals as d
left join customers as c on d.cust_id = c.id
where d.cost in (select max(cost) from daily_totals);