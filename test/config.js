const config = {
  IMPORT_EXCEL: {
    TASK: {
      LIMIT: 2000,
      RULE: [
        'content',
        'parentTask',
        'note',
        'priority',
        'executor',
        'startDate',
        'dueDate',
        'stage'
      ]
    }
  }
}

module.exports = config
