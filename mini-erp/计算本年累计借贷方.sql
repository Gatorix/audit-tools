select "科目代码", "科目名称" ,round(sum("本币借方"),2)as "本年累计借方",round(sum("本币贷方"),2)as "本年累计贷方"  from "科目明细账" 
GROUP by "科目代码"
ORDER by "科目代码"

