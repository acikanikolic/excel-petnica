class UndoRedoManager:
    def __init__(self):
        self.undo_stack = []
        self.redo_stack = []

    def push(self, state):
        self.undo_stack.append(state)
        self.redo_stack.clear()

    def undo(self):
        if len(self.undo_stack) != 0:
            state = self.undo_stack.pop()
            self.redo_stack.append(state)
            return state
        return {}

    def redo(self):
        if len(self.redo_stack) != 0:
            state = self.redo_stack.pop()
            self.undo_stack.append(state)
            return state
        return {}