import { Global, Module } from '@nestjs/common';
import { UserStorageService } from './user-storage.service';

@Global()
@Module({
  providers: [UserStorageService],
  exports: [UserStorageService],
})
export class StorageModule {}
