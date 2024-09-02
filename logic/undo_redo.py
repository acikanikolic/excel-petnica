class UndoRedoManager:
    def __init__(self):
        self.undo_stack = [] #puni stack sa undo
        self.redo_stack = [] #puni stack saa redo

    def push(self, state):
        self.undo_stack.append(state)
        self.redo_stack.clear()

    def undo(self):
        if self.undo_stack:
            state = self.undo_stack.pop()
            self.redo_stack.append(state)
            return self.undo_stack[-1] if self.undo_stack else {}
        return {}

    def redo(self):
        if self.redo_stack:
            state = self.redo_stack.pop()
            self.undo_stack.append(state)
            return state
        return {}
