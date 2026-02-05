from typing import Optional
from PySide6 import QtCore, QtWidgets


class ComboBoxDelegate(QtWidgets.QStyledItemDelegate):
    def __init__(self, items_provider, parent: Optional[QtWidgets.QWidget] = None):
        super().__init__(parent)
        self.items_provider = items_provider

    def createEditor(self, parent, option, index):
        combo = QtWidgets.QComboBox(parent)
        items = self.items_provider() if self.items_provider else []
        combo.addItems(items)
        combo.setEditable(False)
        return combo

    def setEditorData(self, editor, index):
        value = index.model().data(index, QtCore.Qt.ItemDataRole.EditRole) or ""
        editor.setCurrentText(str(value))

    def setModelData(self, editor, model, index):
        model.setData(index, editor.currentText(), QtCore.Qt.ItemDataRole.EditRole)


class SpinBoxDelegate(QtWidgets.QStyledItemDelegate):
    def __init__(
        self,
        minimum: int = 0,
        maximum: int = 9999,
        parent: Optional[QtWidgets.QWidget] = None,
    ):
        super().__init__(parent)
        self.minimum = minimum
        self.maximum = maximum

    def createEditor(self, parent, option, index):
        spin = QtWidgets.QSpinBox(parent)
        spin.setRange(self.minimum, self.maximum)
        return spin

    def setEditorData(self, editor, index):
        value = index.model().data(index, QtCore.Qt.ItemDataRole.EditRole)
        try:
            editor.setValue(int(value))
        except (TypeError, ValueError):
            editor.setValue(self.minimum)

    def setModelData(self, editor, model, index):
        model.setData(index, str(editor.value()), QtCore.Qt.ItemDataRole.EditRole)
