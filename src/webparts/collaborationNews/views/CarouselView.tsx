import * as React from "react";
import Slider from "react-slick";
import styles from "./CarouselView.module.scss";
import "slick-carousel/slick/slick.css";
import "slick-carousel/slick/slick-theme.css";

interface ArrowProps {
  onClick?: () => void;
}

const PrevArrow: React.FC<ArrowProps> = ({ onClick }) => (
  <div className={styles.customPrevArrow} onClick={onClick}>
    <span>◄</span>
  </div>
);

const NextArrow: React.FC<ArrowProps> = ({ onClick }) => (
  <div className={styles.customNextArrow} onClick={onClick}>
    <span>►</span>
  </div>
);

interface CarouselViewProps {
  newsItems: any[];
}

export const CarouselView: React.FC<CarouselViewProps> = ({ newsItems }) => {
  return (
    <div className={styles.carouselContainer}>
      <Slider
        dots={true}
        infinite={true}
        speed={600}
        slidesToShow={1}
        slidesToScroll={1}
        autoplay={true}
        autoplaySpeed={4000}
        arrows={true}
        nextArrow={<NextArrow />}
        prevArrow={<PrevArrow />}
      >
        {newsItems.map((news) => (
          <div key={news.id} className={styles.carouselItem}>
            <img
              src={news.imageUrl}
              alt={news.title}
              className={styles.carouselImage}
              onError={(e) => (e.currentTarget.src = "https://via.placeholder.com/600x300")}
            />
            <div className={styles.carouselContent}>
              <h3>{news.title}</h3>
            </div>
          </div>
        ))} 
      </Slider>
    </div>
  );
};