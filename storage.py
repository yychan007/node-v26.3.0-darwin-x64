import io
import os
import tempfile
from contextlib import contextmanager
from pathlib import Path

_r2_client = None


def r2_enabled():
    required = (
        "R2_ACCOUNT_ID",
        "R2_ACCESS_KEY_ID",
        "R2_SECRET_ACCESS_KEY",
        "R2_BUCKET_NAME",
    )
    return all(os.environ.get(key, "").strip() for key in required)


def storage_backend_name():
    return "r2" if r2_enabled() else "local"


def document_key(stored_filename):
    prefix = os.environ.get("R2_KEY_PREFIX", "documents").strip("/")
    return f"{prefix}/{stored_filename}" if prefix else stored_filename


def get_r2_client():
    global _r2_client
    if _r2_client is None:
        import boto3

        _r2_client = boto3.client(
            "s3",
            endpoint_url=(
                f"https://{os.environ['R2_ACCOUNT_ID']}.r2.cloudflarestorage.com"
            ),
            aws_access_key_id=os.environ["R2_ACCESS_KEY_ID"],
            aws_secret_access_key=os.environ["R2_SECRET_ACCESS_KEY"],
            region_name="auto",
        )
    return _r2_client


def document_exists(stored_filename, local_folder):
    if r2_enabled():
        try:
            from botocore.exceptions import ClientError

            get_r2_client().head_object(
                Bucket=os.environ["R2_BUCKET_NAME"],
                Key=document_key(stored_filename),
            )
            return True
        except ClientError:
            return False

    return (Path(local_folder) / stored_filename).exists()


@contextmanager
def open_document_local_path(stored_filename, local_folder):
    if r2_enabled():
        suffix = Path(stored_filename).suffix or ".bin"
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
        tmp_path = tmp.name
        tmp.close()
        try:
            get_r2_client().download_file(
                os.environ["R2_BUCKET_NAME"],
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
    size = temp.stat().st_size

    if r2_enabled():
        get_r2_client().upload_file(
            str(temp),
            os.environ["R2_BUCKET_NAME"],
            document_key(stored_filename),
        )
        temp.unlink(missing_ok=True)
        return size

    dest = Path(local_folder) / stored_filename
    dest.parent.mkdir(parents=True, exist_ok=True)
    temp.replace(dest)
    return size


def delete_document(stored_filename, local_folder):
    if r2_enabled():
        get_r2_client().delete_object(
            Bucket=os.environ["R2_BUCKET_NAME"],
            Key=document_key(stored_filename),
        )
        return

    path = Path(local_folder) / stored_filename
    if path.exists():
        path.unlink()


def read_document_bytes(stored_filename, local_folder):
    if r2_enabled():
        buffer = io.BytesIO()
        get_r2_client().download_fileobj(
            os.environ["R2_BUCKET_NAME"],
            document_key(stored_filename),
            buffer,
        )
        buffer.seek(0)
        return buffer.read()

    return (Path(local_folder) / stored_filename).read_bytes()
