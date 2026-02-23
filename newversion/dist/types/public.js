/**
 * Public type contracts (JSDoc-only for JS projects).
 */

/** @typedef {{id:string,title:string,parentId:string|null,childCount:number,createdAt?:string,updatedAt?:string}} SpaceNode */
/** @typedef {{id:string,spaceId:string,parentId:string,templateId:string,title:string,index?:string,childCount:number,createdAt?:string,updatedAt?:string}} JournalNode */

export const VERSION = '2.0.0';

export const BACKUP_FORMAT = 'sdo-backup';
export const ENCRYPTED_BACKUP_FORMAT = 'sdo-backup-encrypted';
export const SIGNED_BACKUP_FORMAT = 'sdo-backup-signed';
export const DELTA_BACKUP_FORMAT = 'sdo-backup-delta';
