-- Adapting Cisco SQL select statements for more focused and performant queries
-- in T-SQL
select
    at.EnterpriseName as AgentTeam,
    p.LastName + ', ' + p.FirstName as AgentName,
    dateadd(hh,12,dateadd(dd, DATEDIFF(dd, 0, asg.DateTime), 0)) as Date,
    sum(isnull(asg.CallsHandled, 0)) as InboundHandledCalls,
    isnull(sum(isnull(asg.HandledCallsTime, 0)) / sum(isnull(asg.CallsHandled, 0)), 0) as InboundAHT,
    sum(isnull(ai.LoggedOnTime, 0)) as LogonDuration,
    case when sum(isnull(ai.LoggedOnTime, 0)) = 0.0 then 0.0
        else 1.0 * sum(isnull(ai.NotReadyTime, 0)) / sum(isnull(ai.LoggedOnTime, 0))
        end as PctNotReady,
    case when sum(isnull(ai.LoggedOnTime, 0)) = 0.0 then 0.0
        else 1.0 * (sum(isnull(asg.WorkNotReadyTime, 0)) + 
            sum(isnull(asg.WorkReadyTime, 0))) / 
            sum(isnull(ai.LoggedOnTime, 0))
        end as PctWrap,
    isnull(sum(isnull(asg.AnswerWaitTime, 0)) / sum(isnull(asg.CallsAnswered, 0)), 0) as ASA,
    sum(isnull(asg.HoldTime, 0)) as HoldTime
from (
    select
        asgi.SkillTargetID as AgentSkillTargetID,
        asgi.DateTime as DateTime,
        sum(isnull(asgi.AnswerWaitTime, 0)) as AnswerWaitTime,
        sum(isnull(asgi.CallsAnswered, 0)) as CallsAnswered,
        sum(isnull(asgi.CallsHandled, 0)) as CallsHandled,
        sum(isnull(asgi.HandledCallsTalkTime, 0)) as HandledCallsTalkTime,
        sum(isnull(asgi.HoldTime, 0)) as HoldTime,
        sum(isnull(asgi.WorkNotReadyTime, 0)) as WorkNotReadyTime,
        sum(isnull(asgi.WorkReadyTime, 0)) as WorkReadyTime
    from Agent_Skill_Group_Interval asgi with(nolock)
    group by asgi.SkillTargetID, asgi.DateTime
    ) as asg
    join Agent_Interval ai with(nolock)
                    on asg.AgentSkillTargetID = ai.SkillTargetID
                    and asg.DateTime = ai.DateTime
    join Agent_Team_Member atm with (nolock)
                    on asg.AgentSkillTargetID = atm.SkillTargetID
    join Agent_Team at with (nolock)
                    on atm.AgentTeamID = at.AgentTeamID
    join Agent a with (nolock)
                    on asg.AgentSkillTargetID = a.SkillTargetID
    join Person p with (nolock)
                    on a.PersonID = p.PersonID
where (DATEPART(dw, asg.DateTime) in(2,3,4,5,6,7,1) and 
    asg.DateTime between '2022-01-01 00:07:30' and '2022-01-31 17:00:00' 
    and convert([char], asg.DateTime, 108) between '07:23:00' and '17:07:59') 
    and (at.AgentTeamID IN (5015)) and
    1 = 1
group by
    at.EnterpriseName,
    p.LastName + ', ' + p.FirstName,
    dateadd(day, DATEDIFF(dd, 0, asg.DateTime), 0)
order by
    at.EnterpriseName,
    p.LastName + ', ' + p.FirstName,
    dateadd(day, DATEDIFF(dd, 0, asg.DateTime), 0);