from collections.abc import Generator
from typing import Any
import tempfile
import io
import os
import subprocess
import requests
from uuid import uuid4
import subprocess
from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage


class Test1Tool(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        print("=== 开始执行Docx转PDF工具 ===", flush=True)

 # 1) 获取文件参数（你之前确认字段名是 'completion_report'）
        print("Step 1/5: 开始获取文件参数...", flush=True)
        input_file = tool_parameters.get('completion_report')
        if isinstance(input_file, list):  # 修正：有可能是 list
            input_file = input_file[0]
        if input_file is None:
            msg = "未获取到文件参数，请检查是否传递了 query"
            print("Step 1/5: 失败 -", msg, flush=True)
            yield self.create_json_message({"status": "error", "message": msg})
            return

        # 2) 获取 url —— File 对象没有 get，要用属性
        file_url = getattr(input_file, "url", None)
        if not file_url:
            msg = "文件参数中缺少 url 属性"
            print("Step 2/5: 失败 -", msg, flush=True)
            yield self.create_json_message({"status": "error", "message": msg})
            return
        print("Step 2/5: 成功获取文件 url ->", file_url, flush=True)

        # 3) 拼接完整 URL
        base = os.environ.get("URL", "http://116.205.179.223")
        full_url = base.rstrip('/') + '/' + file_url.lstrip('/')  # 修正变量名
        print("Step 3/5: 拼接完整 URL ->", full_url, flush=True)

        # 3) 转换
        try:
            result_bytes_io = self.convert_docx_to_pdf(full_url)  # 修正变量名
            result_file_bytes = result_bytes_io.getvalue()
            print(f"Step 3/5: 转换完成 - PDF 大小: {len(result_file_bytes)} bytes", flush=True)
        except subprocess.CalledProcessError as e:
            stderr = (e.stderr.decode(errors='ignore') if hasattr(e, 'stderr') and e.stderr else str(e))
            msg = f"LibreOffice 转换失败: {stderr}"
            print("转换失败 -", msg, flush=True)
            yield self.create_json_message({"status": "error", "message": msg})
            return
        except Exception as e:
            print("转换失败 -", str(e), flush=True)
            yield self.create_json_message({"status": "error", "message": str(e)})
            return

        # 4) 输出文件名
        output_filename = tool_parameters.get("output_filename", "converted.pdf")
        if not output_filename.lower().endswith('.pdf'):
            output_filename += '.pdf'
        print("Step 4/5: 输出文件名 ->", output_filename, flush=True)

        # 5) 生成输出：先返回一条简短的文本/JSON提示（不要留空 text），再上传 BLOB
        try:
            # 给下游一个非空的 text，防止某些提取器去读空文本
            yield self.create_json_message({"text": f"已生成文件: {output_filename}"})

            # 关键：把 PDF 二进制上交给 Dify（Dify 会为我们生成 files 条目与 URL）
            yield self.create_blob_message(
                blob=result_file_bytes,
                meta=self.get_meta_data(
                    mime_type="application/pdf",
                    output_filename=output_filename,
                ),
            )

            # 一定要 return 结束生成器
            print("Step 5/5: 输出消息生成完成，转换流程全部结束", flush=True)
            return
        except Exception as e:
            print("发送输出消息失败 -", str(e), flush=True)
            yield self.create_json_message({"status": "error", "message": str(e)})
            return

    def convert_docx_to_pdf(self, url: str) -> io.BytesIO:
        """
        下载 docx -> 使用 libreoffice 转成 pdf -> 返回 BytesIO
        这里使用独立的 LibreOffice profile（--env:UserInstallation）并设置 timeout。
        """
        print("  3.1: 开始创建临时目录...", flush=True)
        with tempfile.TemporaryDirectory() as tmpdir:
            print(f"  3.1: 临时目录 {tmpdir}", flush=True)

            # 下载文件（带超时）
            print("  3.2: 开始下载文件...", flush=True)
            resp = requests.get(url, timeout=30)
            resp.raise_for_status()
            docx_path = os.path.join(tmpdir, "input.docx")
            with open(docx_path, "wb") as f:
                f.write(resp.content)
            print(f"  3.2: 下载并保存到 {docx_path}, 大小 {len(resp.content)} bytes", flush=True)

            # 为 LibreOffice 使用单独的用户配置目录，避免锁冲突
            lo_user_dir = os.path.join(tmpdir, "lo_user")
            os.makedirs(lo_user_dir, exist_ok=True)

            # LibreOffice 命令
            cmd = [
                "libreoffice",
                "--headless",
                "--invisible",
                "--nologo",
                "--nodefault",
                "--norestore",
                f"-env:UserInstallation=file://{lo_user_dir}",
                "--convert-to", "pdf:writer_pdf_Export",
                docx_path,
                "--outdir", tmpdir
            ]

            print("  3.3: 运行 LibreOffice 命令:", " ".join(cmd), flush=True)

            # 执行命令并捕获 stdout/stderr
            result = subprocess.run(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                timeout=120
            )

            print("LibreOffice STDOUT:", result.stdout.strip())
            print("LibreOffice STDERR:", result.stderr.strip())

            if result.returncode != 0:
                raise RuntimeError(
                    f"LibreOffice 转换失败，错误码 {result.returncode}\n"
                    f"STDOUT: {result.stdout.strip()}\n"
                    f"STDERR: {result.stderr.strip()}"
                )

            # LibreOffice 通常会在同目录下生成 input.pdf（或同名 pdf）
            pdf_path = os.path.join(tmpdir, "input.pdf")
            if not os.path.exists(pdf_path):
                candidates = [f for f in os.listdir(tmpdir) if f.lower().endswith(".pdf")]
                if candidates:
                    pdf_path = os.path.join(tmpdir, candidates[0])
                else:
                    raise FileNotFoundError(f"转换后未找到 PDF 文件，目录内容: {os.listdir(tmpdir)}")

            with open(pdf_path, "rb") as f:
                data = f.read()
            print(f"  3.4: 读取到 PDF {pdf_path}, 大小 {len(data)} bytes", flush=True)

            return io.BytesIO(data)

    def get_meta_data(self, mime_type, output_filename):
        return {
            "mime_type": mime_type,
            "filename": output_filename,
        }
