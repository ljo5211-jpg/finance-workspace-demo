"""Only input locations are searched; generated/editable drafts are excluded."""
def input_workbooks(root):
    for folder in (root, root/'적재'/'원천자료', root/'적재'/'최종본'):
        if not folder.is_dir():
            continue
        for path in sorted(folder.iterdir(), key=lambda p:p.name.lower()):
            if path.is_file() and path.suffix.lower()=='.xlsx' and not path.name.startswith('~$'):
                yield path
