#!/usr/bin/env bash
# 演示三条客服问题。需 mewhelp-mysql 容器 + make seed + make dev 就绪。
set -euo pipefail
BASE=http://localhost:8000

echo "== 1 订单物流(应选中 query_logistics 并按返回结果回答)=="
curl -s $BASE/api/agent -H 'Content-Type: application/json' \
  -d '{"user_id":"demo","message":"订单 1001 的物流到哪了"}' | python3 -m json.tool

echo ""
echo "== 2 退货政策(query_faq 命中)=="
curl -s $BASE/api/agent -H 'Content-Type: application/json' \
  -d '{"user_id":"demo","message":"退货政策是什么"}' | python3 -m json.tool

echo ""
echo "== 3 邮费是多少(答案在「运费怎么算」,关键词「邮费」对不上问句,这是关键词检索的漏召回)=="
curl -s $BASE/api/agent -H 'Content-Type: application/json' \
  -d '{"user_id":"demo","message":"邮费是多少"}' | python3 -m json.tool
