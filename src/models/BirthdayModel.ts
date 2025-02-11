import { BirthdayService } from '../services/BirthdayService';
import { Birthday } from '../types';

export class BirthdayModel {
    private birthdayService: BirthdayService;    

    constructor() {
        this.birthdayService = new BirthdayService();
    }

    public async getTodaysBirthdays(): Promise<Birthday[]> {
        const today = new Date();
        const birthdays = await this.birthdayService.fetchBirthdays();
        return birthdays.filter(birthday => 
            birthday.date.getDate() === today.getDate() &&
            birthday.date.getMonth() === today.getMonth()
        );
    }
}