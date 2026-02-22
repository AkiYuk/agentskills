#!/bin/bash
# skills/<スキル名>/ を .skill ファイル（ZIP）にパッケージングするスクリプト
# 使い方:
#   ./build-skill.sh <スキル名>       # 指定スキルをビルド
#   ./build-skill.sh                  # skills/ 配下の全スキルをビルド

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
SKILLS_DIR="${SCRIPT_DIR}/skills"
OUTPUT_DIR="${SCRIPT_DIR}/output"

build_skill() {
  local name="$1"
  local skill_dir="${SKILLS_DIR}/${name}"
  local output_file="${OUTPUT_DIR}/${name}.skill"

  if [ ! -d "${skill_dir}" ]; then
    echo "エラー: ${skill_dir} が見つかりません" >&2
    return 1
  fi

  mkdir -p "${OUTPUT_DIR}"
  rm -f "${output_file}"

  cd "${skill_dir}"
  zip -r "${output_file}" . \
    -x "*/__pycache__/*" \
    -x "*.pyc"

  echo "生成完了: ${output_file}"
}

if [ $# -ge 1 ]; then
  # 引数指定: 指定スキルをビルド
  build_skill "$1"
else
  # 引数なし: 全スキルをビルド
  if [ ! -d "${SKILLS_DIR}" ]; then
    echo "エラー: ${SKILLS_DIR} が見つかりません" >&2
    exit 1
  fi

  found=0
  for skill_dir in "${SKILLS_DIR}"/*/; do
    [ -d "${skill_dir}" ] || continue
    name="$(basename "${skill_dir}")"
    # _始まりはテンプレート等なのでスキップ
    [[ "${name}" == _* ]] && continue
    build_skill "${name}"
    found=1
  done

  if [ "${found}" -eq 0 ]; then
    echo "エラー: ${SKILLS_DIR} にスキルが見つかりません" >&2
    exit 1
  fi
fi
