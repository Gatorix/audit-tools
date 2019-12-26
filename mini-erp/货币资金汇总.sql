SELECT "核算单位名称",
         "子目名称",
         "科目名称",
         "子目代码",
         "币别",
         "期初金额",
         ifnull(round(sum("借方"),2),0.0) AS "借方合计",
         ifnull(round(sum("贷方"),2),0.0) AS "贷方合计",
         round("期初金额"+ifnull(round(sum("借方"),2),0.0)-ifnull(round(sum("贷方"),2),0.0),2)as"期末余额"
FROM '现金银行日记账1-10月'
WHERE "科目名称"="银行存款" AND "币别"!="无单位"
GROUP BY  "子目代码","币别"
UNION
SELECT "核算单位名称",
         "子目名称",
         "科目名称",
         "子目代码",
         "币别",
         "期初金额",
         ifnull(round(sum("借方"),2),0.0) AS "借方合计",
         ifnull(round(sum("贷方"),2),0.0) AS "贷方合计",
         round("期初金额"+ifnull(round(sum("借方"),2),0.0)-ifnull(round(sum("贷方"),2),0.0),2)as"期末余额"
FROM '现金银行日记账1-10月'
WHERE "科目名称"="其他货币资金" AND "币别"!="无单位"
GROUP BY  "子目代码","币别"
ORDER by "子目代码"