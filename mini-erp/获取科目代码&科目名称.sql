CREATE table "科目余额表2" as 
select "科目代码", "科目名称" ,
round(ifnull(sum("本币借方"),0),2)as "本年累计借方",
round(ifnull(sum("本币贷方"),0),2)as "本年累计贷方",
round(ifnull(sum("本币借方"),0)-ifnull(sum("本币贷方"),0),2) as "期末余额"

from "科目明细账" 
GROUP by "科目代码"
ORDER by "科目代码"