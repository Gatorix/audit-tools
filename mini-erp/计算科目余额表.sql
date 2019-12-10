CREATE TABLE "科目余额表11" AS
SELECT 
"期初余额"."科目代码",
"期初余额"."科目名称",
"期初余额",
ifnull(round(sum("本币借方"),2),0)as "本年累计借方",
ifnull(round(sum("本币贷方"),2),0)as "本年累计贷方",
round("期初余额"+ifnull(round(sum("本币借方"),2),0)-ifnull(round(sum("本币贷方"),2),0),2)as"期末余额"
FROM "期初余额" LEFT JOIN "科目明细账" ON "期初余额"."科目代码"="科目明细账"."科目代码" 
GROUP by "期初余额"."科目代码"
ORDER by "期初余额"."科目代码"