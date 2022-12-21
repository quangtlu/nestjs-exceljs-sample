import { Controller, Get, Header, Post, Res } from '@nestjs/common';
import { UploadedFile, UseInterceptors } from '@nestjs/common/decorators';
import { FileInterceptor } from '@nestjs/platform-express';
import { Response } from 'express';
import { ExcelService } from './excel.service';

@Controller('excel')
export class ExcelController {
  constructor(private excelService: ExcelService) {}

  @Get('/download')
  @Header('Content-type', 'text/xlsx')
  async downloadReport(@Res() res: Response) {
    let result = await this.excelService.downloadExcel();
    res.download(`${result}`);
  }

  @Post('upload')
  @UseInterceptors(FileInterceptor('file'))
  async uploadFile(@UploadedFile() file: Express.Multer.File) {
    return await this.excelService.readFileExcel(file);
  }
}
