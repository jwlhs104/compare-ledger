class BaseReconciliationLogic {
  reconcile() {
    throw new Error("reconcile() 必須由子類別實作");
  }

}