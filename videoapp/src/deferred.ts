/**
 * Copyright (c) 2016 shogogg <shogo@studofly.net>
 *
 * This software is released under the MIT License.
 * http://opensource.org/licenses/mit-license.php
 */
export class Deferred<T> {
  private readonly promiseToResolve: Promise<T>;

  private resolveFunc: (value: T | PromiseLike<T>) => void;

  private rejectFunc: (reason: any) => void;

  static idCounter = 1;

  private deferredId: number;

  constructor() {
    this.deferredId = Deferred.idCounter;
    Deferred.idCounter += 1;
    this.resolveFunc = () => {};
    this.rejectFunc = () => {};
    this.promiseToResolve = new Promise<T>((resolve, reject) => {
      this.resolveFunc = resolve;
      this.rejectFunc = reject;
    });
  }

  get promise(): Promise<T> {
    return this.promiseToResolve;
  }

  resolve = (value: T | PromiseLike<T>): void => {
    this.resolveFunc(value);
  };

  reject = (reason?: any): void => {
    this.rejectFunc(reason);
  };

  get id(): number {
    return this.deferredId;
  }
}
