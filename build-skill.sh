#!/bin/bash
# skills/slide-generator/ を .skill ファイル（ZIP）にパッケージングするスクリプト

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
SKILL_NAME="slide-generator"
SKILL_DIR="${SCRIPT_DIR}/skills/${SKILL_NAME}"
OUTPUT_DIR="${SCRIPT_DIR}/output"
OUTPUT_FILE="${OUTPUT_DIR}/${SKILL_NAME}.skill"

# skills/slide-generator/ の存在チェック
if [ ! -d "${SKILL_DIR}" ]; then
  echo "エラー: ${SKILL_DIR} が見つかりません" >&2
  exit 1
fi

# 出力ディレクトリ作成
mkdir -p "${OUTPUT_DIR}"

# 既存ファイルがあれば削除
rm -f "${OUTPUT_FILE}"

# __pycache__ と .pyc を除外してZIP化
cd "${SKILL_DIR}"
zip -r "${OUTPUT_FILE}" . \
  -x "*/__pycache__/*" \
  -x "*.pyc"

echo "生成完了: ${OUTPUT_FILE}"
