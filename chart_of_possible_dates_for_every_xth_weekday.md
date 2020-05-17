# Chart the possible dates for every xth weekday

## In Excel

There are seven weekdays and [no more than five of any given weekday](https://math.stackexchange.com/a/1547547/80230) in a month. Create a table for the Catresian product of seven weekdays and five weeks.

```
Weekday  WeekNumber
Sunday   1
Monday   1
Tuesday  1
. . .
Saturday 1
Sunday   2
Monday   3
. . . and so forth until
Saturday 5
```

Create seven date columns, one for each date the first Sunday of a month might fall on. Enter dates 1–7 in this row.

### Fun part: The Excel Formula

```excel
=(MOD([The cell above this one],7)+1)(7*[WeekNumber])-7
```

In detail:

*TODO*

## Context Notes
I needed to create a series of monthly events in Outlook that fell on weekdays and avoided certain date ranges. I have not had great success with Outlook's event import functions, and my previous solution for this problem (Magneto Calendar) no longer exists. The solution will require either manually creating events or recurring events.

Outlook has limited options for recurring events; we can recur by date and create events on weekends, or we can recur by xth weekday of the month and potnentially create events on disallowed dates. If I could identify a xth weekday of the month that avoided disallowed dates
