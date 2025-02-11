import { BirthdayModel } from '../models/BirthdayModel';
import { Birthday } from '../types';

export class BirthdayController {
    public model: BirthdayModel;

    constructor(model: BirthdayModel) {
        this.model = model;
    }

    public async getTodaysBirthdays(): Promise<Birthday[]> {
        return await this.model.getTodaysBirthdays();
    }
    
}