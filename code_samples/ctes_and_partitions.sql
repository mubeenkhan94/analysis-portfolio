-- Using dates, CTEs, and partitions to determining the number of users who 
-- make new purchases due to a marketing campaign.

-- we want to see whether a user has made a purchase of products on more 
-- than one date AND that there is a date for multiple products
with purchase_dates as (
    select
        user_id,
        -- earliest date for a purchase by a user
        min(created_at) over (partition by user_id) as min_date,
        -- earliest date for a purchase by a user of a product
        min(created_at) over (
        	partition by user_id, product_id
        	) as min_prd_date
    from marketing_campaign
),
worked as (
    select
        user_id,
        -- if the earliest purchase date is equal to the earliest purchase 
        -- date for a product by each product, then that result doesn't 
        -- matter. If they are not equal, then they purchased another product
        -- on a different date
        case when min_date != min_prd_date then 1 else 0 end as worked
    from purchase_dates
)
select count(distinct(user_id)) from worked where worked = 1