-- case when, subqueries, and sorting: Sort rental popularity in specific buckets.

select 
    case 
        when number_of_reviews = 0 then 'New'
        when number_of_reviews >= 1 and number_of_reviews <= 5 then 'Rising'
        when number_of_reviews >= 6 and number_of_reviews <= 15 then 'Trending Up'
        when number_of_reviews >= 16 and number_of_reviews <= 40 then 'Popular'
        else 'Hot' 
            end as popularity_rating,
    min(price) as minimum_price,
    avg(price) as average_price,
    max(price) as maximum_price
from (
    select 
        price, 
        room_type, 
        host_since, 
        zipcode, 
        number_of_reviews
    from airbnb_host_searches
    group by 1,2,3,4,5
) as temp
group by 1;