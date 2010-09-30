if exists (select 1 from sysviews where viewname = 'GuideStatus') then
	drop view GuideStatus
end if;
 

CREATE VIEW
	GuideStatus
AS
SELECT
	  s.statusId
	, w.werkId
	, isnull(ws.status, s.status) as status
FROM
	StatusGuide s
CROSS JOIN GuideWerk w
LEFT JOIN
	WerkStatusNames ws on s.statusId = ws.statusId and w.werkId = ws.werkId
