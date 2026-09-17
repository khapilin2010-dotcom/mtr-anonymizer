"""Distinguish loss of access, unsent work, and corrupt incoming journal entries."""

def feedback(folder, sync):
    errors = [str(value) for value in sync.get('errors', [])]
    pending = int(sync.get('pending', 0))
    if not sync.get('online'):
        title = 'Нет доступа к папке базы'
    elif pending:
        title = 'Не все решения отправлены в базу'
    elif errors:
        title = 'В базе есть записи с ошибками'
    else:
        return dict(ok=True, title='База доступна', short='База доступна. Решения синхронизированы.', details=f'Папка базы: {folder}', signature=None)
    details = f'Папка базы: {folder}\nНеотправленных записей: {pending}\n'
    if errors:
        details += '\nПричина:\n' + '\n'.join(errors[:3])
    details += '\n\nПринятое решение сохранено на этом компьютере. Не удаляйте локальные данные программы.'
    return dict(ok=False, title=title, short=f'{title}. Неотправленных записей: {pending}.',
                details=details, signature=(str(folder), title))
