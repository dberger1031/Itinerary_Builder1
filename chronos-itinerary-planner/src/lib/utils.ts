/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

export function cn(...inputs: any[]) {
  return inputs.filter(Boolean).join(" ");
}
