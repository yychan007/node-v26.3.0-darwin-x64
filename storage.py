import io
import os
import tempfile
import time
from contextlib import contextmanager
from pathlib import Path

_s3_client = None


def r2_enabled():
    required = (
        "R2_ACCOUNT_ID",
        "R2_ACCESS_KEY_ID",
        "R2_SECRET_ACCESS_KEY",
        "R2_BUCKET_NAME",
    )
    return all(os.environ.get(key, "").strip() for key in required)


def supabase_s3_enabled():
    required = (
        "SUPABASE_S3_ENDPOINT",
        "SUPABASE_S3_ACCESS_KEY_ID",
        "SUPABASE_S3_SECRET_ACCESS_KEY",
        "SUPABASE_S3_BUCKET_NAME",
    )
    return all(os.environ.get(key, "").strip() for key in required)


def object_storage_enabled():
    return r2_enabled() or supabase_s3_enabled()


def storage_backend_name():
    if r2_enabled():
        return "r2"
    if supabase_s3_enabled():
        return "supabase"
    return "local"


def bucket_name():
    if r2_enabled():
        return os.environ["R2_BUCKET_NAME"]
    if supabase_s3_enabled():
        return os.environ["SUPABASE_S3_BUCKET_NAME"]
    return ""


def key_prefix():
    if supabase_s3_enabled():
        return os.environ.get("SUPABASE_S3_KEY_PREFIX", "").strip("/")
    return os.environ.get("R2_KEY_PREFIX", "documents").strip("/")


def document_key(stored_filename):
    prefix = key_prefix()
    return f"{prefix}/{stored_filename}" if prefix else stored_filename


def format_storage_error(exc):
    message = str(exc)
    bucket = bucket_name()
    if "NoSuchBucket" in message or "Bucket not found" in message:
        return (
            f"Supabase bucket '{bucket}' was not found. "
            "In Supabase Dashboard go to Storage → New bucket, create a bucket with "
            f"this exact name ('{bucket}'), then try again. "
            "Also verify SUPABASE_S3_BUCKET_NAME on Render matches that name."
        )
    if "AccessDenied" in message or "403" in message:
        return (
            f"Access denied for bucket '{bucket}'. "
            "Regenerate S3 access keys in Supabase → Project Settings → Storage → "
            "S3 Access Keys, then update SUPABASE_S3_ACCESS_KEY_ID and "
            "SUPABASE_S3_SECRET_ACCESS_KEY on Render."
        )
    if "InvalidAccessKeyId" in message or "SignatureDoesNotMatch" in message:
        return (
            "Invalid Supabase S3 credentials. "
            "Check SUPABASE_S3_ACCESS_KEY_ID, SUPABASE_S3_SECRET_ACCESS_KEY, "
            "SUPABASE_S3_ENDPOINT, and SUPABASE_S3_REGION on Render."
        )
    return message


def get_storage_status():
    backend = storage_backend_name()
    if not object_storage_enabled():
        return {
            "ok": True,
            "backend": backend,
            "persistent": False,
            "bucket": "",
            "key_prefix": "",
            "message": "Files are saved on local disk only (lost after Render restart).",
        }

    bucket = bucket_name()
    prefix = key_prefix()
    status = {
        "ok": False,
        "backend": backend,
        "persistent": True,
        "bucket": bucket,
        "key_prefix": prefix or "(root)",
        "endpoint": os.environ.get("SUPABASE_S3_ENDPOINT", "") if supabase_s3_enabled() else "",
        "region": os.environ.get("SUPABASE_S3_REGION", "") if supabase_s3_enabled() else "",
        "message": "",
        "hint": "",
    }

    try:
        from botocore.exceptions import BotoCoreError, ClientError

        client = get_s3_client()
        client.head_bucket(Bucket=bucket)
        status["ok"] = True
        status["message"] = f"Connected to bucket '{bucket}'."
        return status
    except ClientError as exc:
        code = exc.response.get("Error", {}).get("Code", "")
        status["message"] = exc.response.get("Error", {}).get("Message", str(exc))
        if code in {"404", "NoSuchBucket", "NotFound"}:
            status["hint"] = (
                f"Create bucket '{bucket}' in Supabase → Storage → New bucket. "
                "Set SUPABASE_S3_BUCKET_NAME on Render to the same name."
            )
        elif code in {"403", "AccessDenied"}:
            status["hint"] = (
                "Regenerate S3 access keys in Supabase → Project Settings → Storage."
            )
        else:
            status["hint"] = format_storage_error(exc)
        return status
    except (BotoCoreError, RuntimeError) as exc:
        status["message"] = str(exc)
        status["hint"] = format_storage_error(exc)
        return status
    except Exception as exc:
        status["message"] = str(exc)
        status["hint"] = format_storage_error(exc)
        return status


def get_s3_client():
    global _s3_client
    if _s3_client is not None:
        return _s3_client

    try:
        import boto3
    except ImportError as exc:
        raise RuntimeError(
            "Object storage is configured but boto3 is not installed."
        ) from exc

    if r2_enabled():
        _s3_client = boto3.client(
            "s3",
            endpoint_url=(
                f"https://{os.environ['R2_ACCOUNT_ID']}.r2.cloudflarestorage.com"
            ),
            aws_access_key_id=os.environ["R2_ACCESS_KEY_ID"],
            aws_secret_access_key=os.environ["R2_SECRET_ACCESS_KEY"],
            region_name="auto",
        )
        return _s3_client

    if supabase_s3_enabled():
        _s3_client = boto3.client(
            "s3",
            endpoint_url=os.environ["SUPABASE_S3_ENDPOINT"],
            aws_access_key_id=os.environ["SUPABASE_S3_ACCESS_KEY_ID"],
            aws_secret_access_key=os.environ["SUPABASE_S3_SECRET_ACCESS_KEY"],
            region_name=os.environ.get("SUPABASE_S3_REGION", "eu-west-1"),
        )
        return _s3_client

    raise RuntimeError("No object storage backend configured.")


def document_exists(stored_filename, local_folder):
    if object_storage_enabled():
        from botocore.exceptions import BotoCoreError, ClientError

        key = document_key(stored_filename)
        try:
            client = get_s3_client()
            client.head_object(Bucket=bucket_name(), Key=key)
            return True
        except ClientError as exc:
            code = exc.response.get("Error", {}).get("Code", "")
            if code in {"404", "NoSuchKey", "NotFound", "403"}:
                try:
                    client = get_s3_client()
                    client.get_object(Bucket=bucket_name(), Key=key)
                    return True
                except Exception:
                    return False
            print(f"Storage head_object error for {key}: {exc}")
            return False
        except (BotoCoreError, RuntimeError) as exc:
            print(f"Storage check error for {stored_filename}: {exc}")
            return False
        except Exception as exc:
            print(f"Unexpected storage check error for {stored_filename}: {exc}")
            return False

    return (Path(local_folder) / stored_filename).exists()


def list_stored_document_filenames(local_folder):
    if object_storage_enabled():
        from botocore.exceptions import BotoCoreError, ClientError

        prefix = key_prefix()
        list_prefix = f"{prefix}/" if prefix else ""
        filenames = set()
        try:
            client = get_s3_client()
            paginator = client.get_paginator("list_objects_v2")
            paginate_kwargs = {"Bucket": bucket_name()}
            if list_prefix:
                paginate_kwargs["Prefix"] = list_prefix
            for page in paginator.paginate(**paginate_kwargs):
                for item in page.get("Contents", []):
                    key = item.get("Key") or ""
                    if not key or key.endswith("/"):
                        continue
                    if list_prefix and key.startswith(list_prefix):
                        name = key[len(list_prefix) :]
                    else:
                        name = key.rsplit("/", 1)[-1]
                    if name:
                        filenames.add(name)
        except (ClientError, BotoCoreError, RuntimeError) as exc:
            print(f"Storage list error: {exc}")
            return None
        except Exception as exc:
            print(f"Unexpected storage list error: {exc}")
            return None
        return filenames

    folder = Path(local_folder)
    if not folder.exists():
        return set()
    return {path.name for path in folder.iterdir() if path.is_file()}


def verify_document_stored(stored_filename, local_folder):
    if document_exists(stored_filename, local_folder):
        return

    if object_storage_enabled():
        for _ in range(2):
            time.sleep(0.5)
            if document_exists(stored_filename, local_folder):
                return

    backend = storage_backend_name()
    raise RuntimeError(
        "File was not saved to storage. "
        f"Current backend: {backend}. "
        "Configure Supabase S3 (or R2) on Render for persistent PDF storage."
    )


@contextmanager
def open_document_local_path(stored_filename, local_folder):
    if object_storage_enabled():
        suffix = Path(stored_filename).suffix or ".bin"
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
        tmp_path = tmp.name
        tmp.close()
        try:
            get_s3_client().download_file(
                bucket_name(),
                document_key(stored_filename),
                tmp_path,
            )
            yield tmp_path
        finally:
            Path(tmp_path).unlink(missing_ok=True)
    else:
        yield str(Path(local_folder) / stored_filename)


def save_document(stored_filename, temp_path, local_folder):
    temp = Path(temp_path)
    if not temp.exists():
        raise FileNotFoundError(f"Temporary upload file not found: {temp}")

    size = temp.stat().st_size

    if object_storage_enabled():
        try:
            get_s3_client().upload_file(
                str(temp),
                bucket_name(),
                document_key(stored_filename),
            )
        except Exception as exc:
            raise RuntimeError(format_storage_error(exc)) from exc
        temp.unlink(missing_ok=True)
        return size

    dest = Path(local_folder) / stored_filename
    try:
        dest.parent.mkdir(parents=True, exist_ok=True)
        temp.replace(dest)
    except Exception as exc:
        raise RuntimeError(f"Local storage save failed: {exc}") from exc
    return size


def delete_document(stored_filename, local_folder):
    if object_storage_enabled():
        try:
            get_s3_client().delete_object(
                Bucket=bucket_name(),
                Key=document_key(stored_filename),
            )
        except Exception as exc:
            print(f"Storage delete error for {stored_filename}: {exc}")
        return

    path = Path(local_folder) / stored_filename
    if path.exists():
        path.unlink()


def read_document_bytes(stored_filename, local_folder):
    if object_storage_enabled():
        buffer = io.BytesIO()
        get_s3_client().download_fileobj(
            bucket_name(),
            document_key(stored_filename),
            buffer,
        )
        buffer.seek(0)
        return buffer.read()

    return (Path(local_folder) / stored_filename).read_bytes()
